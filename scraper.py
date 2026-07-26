from playwright.sync_api import sync_playwright
import time
import openpyxl
import os

def find_businesses_without_website(district, query_term):
    search_term = f"{district} {query_term}"
    
    # Excel dosya adını başta hazırlayalım ki anlık kayıt yapabilelim
    safe_district = district.replace(" ", "_").lower()
    safe_term = query_term.replace(" ", "_").lower()
    filename = f"detayli_analiz_{safe_district}_{safe_term}.xlsx"
    
    # Eğer dosya daha önce kalmışsa temizleyelim ve başlıkları atalım
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Hedef İşletmeler"
    ws.append(["Semt", "Kategori", "İşletme Adı", "Telefon", "E-posta", "Sosyal Medya / Web", "Adres", "Google Maps Linki"])
    wb.save(filename)

    with sync_playwright() as p:
        # Tarayıcımızı açıyoruz (İstersen headless=True yaparak arka planda da çalıştırabilirsin)
        browser = p.chromium.launch(headless=False) 
        page = browser.new_page(viewport={"width": 1280, "height": 800})
        
        print(f"\n--- [PRO MOD] Bölge: {district} | Sektör: {query_term} ---")
        print("Google Haritalar'a bağlanılıyor...")
        page.goto("https://www.google.com/maps")
        
        # 1. Çerez Onayı
        try:
            page.wait_for_timeout(2000) 
            cookie_button = page.locator('button:has-text("Tümünü kabul et"), button:has-text("Accept all"), button:has-text("Kabul ediyorum")')
            if cookie_button.count() > 0 and cookie_button.first.is_visible():
                cookie_button.first.click()
                print("Çerez onayı geçildi.")
                page.wait_for_timeout(1000)
        except:
            pass 
        
        # 2. Arama Kutusuna Yazma
        try:
            search_box = page.locator('input#searchboxinput, input[name="q"], input.searchboxinput')
            search_box.first.wait_for(state="visible", timeout=15000) 
            search_box.first.fill(search_term)
            search_box.first.press('Enter')
            print(f"'{search_term}' araması başlatıldı, sonuçlar listeleniyor...")
            page.wait_for_timeout(5000) 
        except Exception as e:
            print("Arama kutusu bulunamadı, sistem kapatılıyor.")
            browser.close()
            return
        
        # 3. Listeyi Sonuna Kadar Kaydırma (Derinlemesine Tarama)
        listings_locator = page.locator('a[href*="https://www.google.com/maps/place/"]')
        previously_counted = 0
        scrolling_attempts = 0
        
        print("Listenin sonuna ulaşmak için kaydırma işlemi yapılıyor...")
        while True:
            count = listings_locator.count()
            if count == 0:
                break
                
            listings_locator.nth(count - 1).scroll_into_view_if_needed()
            page.wait_for_timeout(2500) 
            
            new_count = listings_locator.count()
            if new_count == previously_counted:
                scrolling_attempts += 1
                if scrolling_attempts >= 3: 
                    break
            else:
                scrolling_attempts = 0
                previously_counted = new_count
                
            print(f"Şu ana kadar {new_count} işletme yüklendi...")
            
        # 4. Linkleri Toplama
        total_places = listings_locator.count()
        print(f"\nKaydırma bitti. Toplam {total_places} işletme incelenmek üzere kuyruğa alındı.")
        
        place_urls = []
        for i in range(total_places):
            href = listings_locator.nth(i).get_attribute("href")
            if href:
                place_urls.append(href)
                
        found_count = 0

        # 5. İşletmeleri Tek Tek Ziyaret Etme ve Anlık Kayıt
        for index, url in enumerate(place_urls):
            try:
                page.goto(url)
                page.wait_for_timeout(2500) 
                
                # İşletme Adı
                business_name_element = page.locator('h1')
                business_name = business_name_element.first.inner_text() if business_name_element.count() > 0 else "İsimsiz İşletme"
                
                # Web Sitesi Var mı?
                has_website = page.locator('a[data-item-id="authority"]').is_visible()
                
                if not has_website:
                    found_count += 1
                    print(f"[{index + 1}/{len(place_urls)}] 🎯 YAKALANDI (Web sitesi yok): {business_name}")
                    
                    # Telefon Numarası
                    phone = "Bulunamadı"
                    phone_el = page.locator('button[data-item-id^="phone:tel:"]')
                    if phone_el.count() > 0:
                        phone = phone_el.first.get_attribute("aria-label")
                        phone = phone.replace("Telefon: ", "").strip() if phone else phone_el.first.inner_text().strip()
                            
                    # Adres
                    address = "Bulunamadı"
                    addr_el = page.locator('button[data-item-id="address"]')
                    if addr_el.count() > 0:
                        address = addr_el.first.get_attribute("aria-label")
                        address = address.replace("Adres: ", "").strip() if address else addr_el.first.inner_text().strip()
                    
                    # Ekstra: Sayfa İçinden E-posta Avcılığı (Eğer varsa)
                    email = "Bulunamadı"
                    try:
                        page_content = page.content()
                        import re
                        emails = re.findall(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+', page_content)
                        # Google'ın kendi servis maillerini eleyelim
                        valid_emails = [e for e in emails if "google" not in e and "gstatic" not in e]
                        if valid_emails:
                            email = valid_emails[0]
                    except:
                        pass

                    # Sosyal Medya veya Diğer Bağlantılar (Instagram, Facebook vb.)
                    social_media = "Bulunamadı"
                    try:
                        social_el = page.locator('a[href*="instagram.com"], a[href*="facebook.com"]')
                        if social_el.count() > 0:
                            social_media = social_el.first.get_attribute("href")
                    except:
                        pass

                    # ANLIK EXCEL'E YAZMA (Veri kaybını önlemek için her bulduğunda dosyaya ekler)
                    current_wb = openpyxl.load_workbook(filename)
                    current_ws = current_wb.active
                    current_ws.append([district, query_term, business_name, phone, email, social_media, address, url])
                    current_wb.save(filename)
                    
                else:
                    print(f"[{index + 1}/{len(place_urls)}] ❌ PAS GEÇİLDİ (Web sitesi var): {business_name}")
                    
            except Exception as e:
                print(f"[{index + 1}/{len(place_urls)}] Bir işletme okunurken hata atlandı.")
                
        browser.close()
        print(f"\n🎉 İŞLEM TAMAMLANDI! Toplam {found_count} adet web sitesiz işletme bulundu.")
        print(f"📂 Veriler güvenle '{filename}' dosyasına kaydedildi.")

if __name__ == "__main__":
    try:
        # BURADAN SEMT VE KELİMEYİ İSTEDİĞİN GİBİ DEĞİŞTİREBİLİRSİN
        hedef_semt = input("Lütfen hedef semti giriniz (örn: Beşiktaş): ").strip()
        hedef_kelime = input("Lütfen hedef kelimeyi giriniz (örn: marangoz): ").strip()

        find_businesses_without_website(hedef_semt, hedef_kelime)
    except KeyboardInterrupt:
        print("\nİşlem kullanıcı tarafından durduruldu.")
        exit(0)
    except Exception as e:
        print(f"Beklenmeyen bir hata oluştu: {e}")
        exit(1)
