from playwright.sync_api import sync_playwright
import pandas as pd
import time
import re
import re


import os


def search_inaproc(keyword, headless=True, min_price=0, max_price=0, location_filter=None, max_pages=1, enable_comparison=False, limit_products=0, sort_order="Paling Sesuai"):
    """
    Scrapes katalog.inaproc.id using Playwright.
    sort_order: "Paling Sesuai", "Harga Terendah", "Harga Tertinggi"
    """
    results = []
    
    # Buat folder screenshot jika belum ada
    if enable_comparison:
        if not os.path.exists("screenshots"):
            os.makedirs("screenshots")
    
    with sync_playwright() as p:
        # Launch browser dengan argumen anti-bot
        browser = p.chromium.launch(
            headless=False,  # Harus False agar tidak terdeteksi bot dengan mudah
            args=["--disable-blink-features=AutomationControlled"] 
        )
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()
        
        # Navigate
        # Base URL
        url = f"https://katalog.inaproc.id/search?keyword={keyword}"
        
        # Add Filters (Price URL)
        if min_price > 0:
            url += f"&minPrice={min_price}"
        if max_price > 0:
            url += f"&maxPrice={max_price}"
            
        print(f"Mengakses: {url}")
        page.goto(url, timeout=60000)
        
        print(f"Mengakses: {url}")
        page.goto(url, timeout=60000)
        
        # --- LOGIKA SORTING (UI AUTOMATION) ---
        if sort_order != "Paling Sesuai":
            print(f"Mencoba sorting: {sort_order}")
            try:
                # Tunggu konten loaded
                page.wait_for_selector("text=Rp", timeout=15000)
                
                # Cari tombol sort dengan berbagai kemungkinan teks
                # Biasanya teks defaultnya "Paling Sesuai"
                # Kita cari elemen teks yang clickable
                sort_btn = None
                
                # Coba 1: Button/Div spesifik dengan teks "Paling Sesuai"
                for txt in ["Paling Sesuai", "Urutkan", "Relevansi"]:
                    candidate = page.locator(f"text={txt}").first
                    if candidate.count() > 0 and candidate.is_visible():
                        # Cek apakah parent/self clickable
                        sort_btn = candidate
                        break
                
                if sort_btn:
                    print(f"Tombol sort ditemukan: {sort_btn.inner_text()}")
                    # Klik elemen tersebut (atau parent jika perlu)
                    sort_btn.click() # Playwright biasanya pintar klik center
                    time.sleep(1.5) # Tunggu dropdown
                    
                    # Cari Opsi
                    # Opsi biasanya teks "Harga Terendah", "Harga Tertinggi"
                    target_opt = page.locator(f"text={sort_order}").first
                    
                    if target_opt.count() > 0 and target_opt.is_visible():
                         print(f"Clicking option: {sort_order}")
                         target_opt.click()
                         # Tunggu efek sorting (page reload atau re-render grid)
                         time.sleep(5)
                    else:
                         print(f"Opsi '{sort_order}' TIDAK muncul di dropdown.")
                         # Debug: print apa yang muncul
                         print("Visible text in dropdown area:")
                         print(page.locator("div[role='listbox'], ul").first.inner_text())
                else:
                    print("Tombol pembuka sort TIDAK ditemukan.")
                    
            except Exception as e:
                print(f"Gagal sorting: {e}")

        # --- LOGIKA FILTER LOKASI (UI AUTOMATION) ---
        if location_filter:
            print(f"Mencoba menerapkan filter lokasi: {location_filter}")
            try:
                # 1. Cari & Expand Accordion 'Lokasi Pengiriman'
                # Gunakan text locator yang specific
                # Tambahkan timeout agar halaman benar-benar ready
                try:
                    page.wait_for_selector("text=Lokasi Pengiriman", timeout=10000)
                except:
                    print("Timeout menunggu accordion lokasi.")

                loc_btn = page.locator("div").filter(has_text="Lokasi Pengiriman").last
                
                if loc_btn.count() > 0:
                    loc_btn.click()
                    time.sleep(2) # Tunggu animasi expand (diperlama sedikit)
                
                # 2. Cari Checkbox
                found = False

                # Strategi 1: Coba cari text langsung
                target_label = page.locator("label").filter(has_text=location_filter).first
                if target_label.count() > 0 and target_label.is_visible():
                    print(f"[OK] Lokasi '{location_filter}' ditemukan langsung. Clicking...")
                    target_label.click()
                    found = True
                else:
                    # Scroll ke bawah sidebar untuk memancing lazy load
                    print("Scrolling sidebar...")
                    try:
                        last_label = page.locator("label.flex.items_center").last
                        if last_label.count() > 0:
                            last_label.scroll_into_view_if_needed()
                            time.sleep(2)
                    except: pass

                # Strategi 2: Coba cari "Lihat Selengkapnya" dan klik satu per satu
                if not found:
                    print(f"[INFO] '{location_filter}' tidak visible. Mencari tombol 'Lihat Selengkapnya'...")
                    
                    try:
                        page.wait_for_selector("text=Lihat Selengkapnya", timeout=5000)
                    except: pass
                    
                    show_more_btns = page.locator("text=Lihat Selengkapnya").all()
                    
                    for idx, btn in enumerate(show_more_btns):
                        try:
                            if btn.is_visible():
                                print(f"Clicking 'Lihat Selengkapnya' #{idx}...")
                                btn.click()
                                time.sleep(2) # Tunggu UI update (increased from 1s)
                                
                                # Cek modal
                                modal = page.locator("div[role='dialog']")
                                if modal.count() > 0 and modal.is_visible():
                                    print("Modal muncul! Menggunakan fitur search modal...")
                                    search_input = modal.locator("input[placeholder*='Cari']")
                                    
                                    # RETRY MECHANISM untuk Search Box
                                    for attempt in range(3):
                                        if search_input.count() > 0:
                                            print(f"Mengisi search box (percobaan {attempt+1})...")
                                            search_input.fill(location_filter)
                                            # Tunggu sebentar agar AJAX search jalan
                                            # Next.js RSC bisa lambat, kasih 6 detik
                                            time.sleep(6) 
                                            
                                            # Cek hasil
                                            # Kita harus spesifik: Cari label yang visible
                                            modal_results = modal.locator("label:visible")
                                            count = modal_results.count()
                                            
                                            found_match = False
                                            if count > 0:
                                                print(f"Ditemukan {count} opsi di modal.")
                                                # Scan text satu per satu
                                                # Playwright .all() bisa stale, pakai nth
                                                for i in range(count):
                                                    txt = modal_results.nth(i).inner_text().strip()
                                                    # print(f"Opsi {i}: {txt}")
                                                    if location_filter.lower() in txt.lower():
                                                         found_match = True
                                                         break
                                                
                                                if found_match:
                                                    break # Sukses, lanjut ke pemilihan
                                            
                                            print("Hasil yang cocok belum muncul, retry waiting...")
                                            time.sleep(3)
                                        else:
                                            time.sleep(1)

                                    # Pilih hasil yang SUDAH divalidasi
                                    modal_results = modal.locator("label:visible").all()
                                    target_result = None
                                    for r in modal_results:
                                        if location_filter.lower() in r.inner_text().lower():
                                            target_result = r
                                            break
                                    
                                    if target_result:
                                        clicked_text = target_result.inner_text().strip()
                                        print(f"Clicking result: '{clicked_text}'")
                                        
                                        # Verifikasi ganda: Jangan klik jika teksnya jauh berbeda (misal cuma 1 huruf sama)
                                        if location_filter.lower() not in clicked_text.lower():
                                            print(f"[BAHAYA] Akan mengklik '{clicked_text}' tapi tidak sesuai '{location_filter}'. ABORT.")
                                            # Skip click
                                        else:
                                            target_result.scroll_into_view_if_needed()
                                            target_result.click()
                                            time.sleep(2)
                                            
                                            # Close modal with retry
                                            save_btns = modal.locator("button").filter(has_text="Terapkan").all()
                                            if len(save_btns) > 0:
                                                save_btns[0].click()
                                            else:
                                                # Close icon fallback
                                                close_btn = modal.locator("button svg").first
                                                if close_btn.count() > 0: close_btn.click()
                                                else: page.mouse.click(0, 0)
                                                
                                            found = True
                                            time.sleep(8) # Loading page setelah filter diterapkan BISA LAMA
                                            break
                                    else:
                                        print("Tidak ada label yang cocok di modal.")
                                        # Close modal to clean up
                                        page.keyboard.press("Escape")
                                        time.sleep(1)
                        except Exception as inner_e:
                            print(f"Gagal interaksi tombol #{idx}: {inner_e}")
                
                if not found:
                    print(f"[WARN] Gagal menemukan filter lokasi '{location_filter}'.")
                    
            except Exception as e:
                print(f"Gagal menerapkan filter lokasi UI: {e}")

        # Variable untuk comparison limit
        comparison_count = 0
        
        # --- MAIN PAGINATION LOOP ---
        for current_page_num in range(1, max_pages + 1):
            print(f"--- Scraping Halaman {current_page_num} ---")
            
            try:
                # Tunggu skeleton loader menghilang atau setidaknya data muncul
                # Kita tunggu elemen yang BUKAN skeleton (animasi pulse)
                # Biasanya produk ada di dalam grid.
                # Strategy: Tunggu ada text harga "Rp"
                print("Menunggu data dimuat...")
                page.wait_for_selector("text=Rp", timeout=15000)
                
                # Scroll down to load images (lazy load)
                # Scroll berkali-kali untuk memastikan infinite scroll (dalam satu halaman) kelar
                # Meskipun kita pakai pagination, dalam satu page mungkin ada 60 item yg perlu load
                for _ in range(5):
                    page.mouse.wheel(0, 2000)
                    time.sleep(1)
                
                # Selector Produk yang VALID (Analisis dari debug_page.html)
                # Produk dibungkus dalam <a> tag yang merupakan anak langsung dari div.grid
                # Class grid di Inaproc: "mt-6 grid grid-cols-1 ..."
                product_cards = page.locator("div.grid > a").all()
                
                print(f"Ditemukan {len(product_cards)} produk di halaman ini.")
                
                inactive_count_on_page = 0
                
                for card in product_cards:
                    try:
                        # --- SMART PAGINATION LOGIC ---
                        # Cek status produk (Active/Inactive)
                        # Gunakan text content card untuk deteksi "Belum Aktif" dsb.
                        card_text = card.inner_text()
                        # Keywords yang diminta user + variasi
                        if "Belum Aktif" in card_text or "Stok Habis" in card_text:
                            inactive_count_on_page += 1
                            continue

                        # Extract Data
                        link = "https://katalog.inaproc.id" + card.get_attribute("href")
                        
                        # Judul (Class: line-clamp-2 text-sm text-tertiary500)
                        title_el = card.locator("div.line-clamp-2")
                        title = title_el.inner_text() if title_el.count() > 0 else "N/A"
                        
                        # Harga (Class: w-fit truncate text-sm font-bold)
                        price_el = card.locator("div.w-fit").first
                        price = price_el.inner_text() if price_el.count() > 0 else "N/A"
                        
                        # Vendor (Ada di dalam div h-4 cursor-pointer, span kedua biasanya nama vendor)
                        # Struktur: Div > Span (Kota) + Span (Vendor)
                        vendor_container = card.locator("div.h-4.cursor-pointer span")
                        
                        location = "N/A"
                        vendor = "N/A"
                        
                        if vendor_container.count() >= 2:
                            location = vendor_container.nth(0).inner_text()
                            vendor = vendor_container.nth(1).inner_text()
                        elif vendor_container.count() == 1:
                            # Fallback basic
                            text = vendor_container.first.inner_text()
                            if "Kab." in text or "Kota" in text:
                                location = text
                            else:
                                vendor = text

                        # --- STRICT POST-FILTERING (PYTHON SIDE) ---
                        # Pastikan lokasi sesuai permintaan user, jika filter UI gagal
                        if location_filter:
                            # Normalisasi text (lowercase)
                            loc_req = location_filter.lower()
                            loc_found = location.lower()
                            # Cek apakah substring ada
                            if loc_req not in loc_found:
                                # Kadang lokasi ada di Vendor name, check juga vendor as fallback (opsional)
                                # if loc_req in vendor.lower(): pass else:
                                # Strict mode: skip
                                # print(f"Skip lokasi tidak sesuai: {location} != {location_filter}")
                                continue
                                
                        # Image
                        img_tag = card.locator("img").first
                        img_url = img_tag.get_attribute("src") if img_tag.count() > 0 else ""
                        
                        screenshot_path = None
                        
                        # --- COMPARISON MODE LOGIC ---
                        if enable_comparison:
                            print(f"[Comparison] Membuka detail produk: {title}...")
                            try:
                                # Open link in new page
                                detail_page = context.new_page()
                                detail_page.goto(link, timeout=45000)
                                detail_page.wait_for_load_state("networkidle")
                                time.sleep(2) # Give explicit render time
                                
                                # Sanitize vendor for filename
                                safe_name = re.sub(r'[\\/*?:"<>|]', "", vendor)[:50].strip()
                                # Fallback jika vendor kosong
                                if not safe_name:
                                    safe_name = re.sub(r'[\\/*?:"<>|]', "", title)[:20].strip()
                                    
                                filename = f"{safe_name}_{int(time.time())}.png"
                                filepath = os.path.join(os.getcwd(), "screenshots", filename)
                                
                                print(f"Saving screenshot to: {filepath}")
                                detail_page.screenshot(path=filepath, full_page=True)
                                detail_page.close()
                                
                                screenshot_path = filepath
                                comparison_count += 1
                                
                            except Exception as detail_e:
                                print(f"Gagal screenshot detail: {detail_e}")
                                if 'detail_page' in locals(): detail_page.close()
                        
                        results.append({
                            "Nama Produk": title,
                            "Harga": price,
                            "Penyedia": vendor,
                            "Lokasi": location,
                            "Link": link,
                            "Gambar": img_url,
                            "Screenshot": screenshot_path # New Field
                        })
                        
                        # Check Comparison Limit
                        if enable_comparison and limit_products > 0 and comparison_count >= limit_products:
                            print(f"Limit comparison {limit_products} tercapai. Berhenti.")
                            break
                            
                    except Exception as e:
                        print(f"Error parsing card: {e}")
                        continue
                        
                # --- SMART STOP CONDITION ---
                if len(product_cards) > 0:
                    inactive_ratio = inactive_count_on_page / len(product_cards)
                    # Jika > 50% produk di halaman ini inactive, kemungkinan halaman selanjutnya sampah.
                    if inactive_ratio > 0.5: 
                        print(f"[SMART STOP] {inactive_count_on_page}/{len(product_cards)} produk tidak aktif. Menghentikan pagination.")
                        break
                
                # Break outer loop if comparison limit reached
                if enable_comparison and limit_products > 0 and comparison_count >= limit_products:
                    break

                # --- NAVIGASI KE HALAMAN BERIKUTNYA ---
                if current_page_num < max_pages:
                    next_page_num = current_page_num + 1
                    print(f"Mencoba pindah ke halaman {next_page_num}...")
                    
                    # Cari tombol dengan angka halaman berikutnya
                    # Kita pakai locator tombol yang text-nya persis angka
                    # Karena kadang ada "10", "11", kita pakai exact match atau regex boundaries
                    
                    # Scroll ke bawah dulu pol
                    page.keyboard.press("End")
                    time.sleep(1)
                    
                    # Locator untuk tombol angka dengan EXACT MATCH regex
                    # ^2$ matches "2", but not "12"
                    next_btn = page.locator("button").filter(has_text=re.compile(r"^" + str(next_page_num) + "$")).first
                    
                    if next_btn.count() > 0 and next_btn.is_visible():
                        print(f"Klik tombol halaman {next_page_num}...")
                        next_btn.scroll_into_view_if_needed()
                        next_btn.click()
                        time.sleep(5) # Tunggu load halaman baru
                    else:
                        print(f"Tombol halaman {next_page_num} tidak ditemukan. Berhenti scraping.")
                        # Coba fallback ke tombol 'Next' chevron jika ada (biasanya tombol terakhir tanpa text angka)
                        # Tapi untuk 1-5, angka harusnya ada.
                        break
                        
            except Exception as e:
                print(f"Terjadi kesalahan saat halaman {current_page_num}: {e}")
                page.screenshot(path=f"error_page_{current_page_num}.png")
                # Jangan break, lanjut loop siapa tahu (meski aneh)
            
        browser.close()
    
    return results

if __name__ == "__main__":
    # Test run
    data = search_inaproc("laptop", headless=True)
    print(f"Hasil: {len(data)} produk")
    if data:
        print(data[0])
