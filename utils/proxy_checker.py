import requests
import threading
import queue
import time
import socket
from concurrent.futures import ThreadPoolExecutor

# Test edilecek URL (Google genellikle iyi bir test hedefidir)
TEST_URL = "https://www.google.com"
TIMEOUT = 10  # Saniye cinsinden zaman aşımı süresi
MAX_THREADS = 50  # Maksimum iş parçacığı sayısı
RESULT_FILE = "utils/proxy_valid.txt"  # Sonuçların yazılacağı dosya

# Proxy listesini oku
def read_proxies(file_path):
    with open(file_path, 'r') as f:
        return [line.strip() for line in f if line.strip()]

# Proxy'nin çalışırlığını kontrol et
def check_proxy(proxy, result_queue):
    protocol = "http"
    if ":" in proxy:
        ip, port = proxy.split(":")
        
        # HTTP protokolü için kontrol et
        proxy_dict = {
            "http": f"http://{proxy}",
            "https": f"http://{proxy}"
        }
        
        try:
            # Bağlantı denemesi
            start_time = time.time()
            response = requests.get(TEST_URL, proxies=proxy_dict, timeout=TIMEOUT)
            elapsed_time = time.time() - start_time
            
            # Başarılı yanıt aldık mı?
            if response.status_code == 200:
                print(f"[+] {proxy} - Çalışıyor! Yanıt süresi: {elapsed_time:.2f}s")
                result_queue.put((proxy, elapsed_time))
                return True
            else:
                print(f"[-] {proxy} - Yanıt kodu: {response.status_code}")
                return False
                
        except requests.exceptions.RequestException as e:
            print(f"[-] {proxy} - Hata: {str(e)}")
            return False
        except Exception as e:
            print(f"[-] {proxy} - Beklenmeyen hata: {str(e)}")
            return False
    return False

# İş parçacığı fonksiyonu
def worker(proxy_queue, result_queue):
    while not proxy_queue.empty():
        try:
            proxy = proxy_queue.get(block=False)
            check_proxy(proxy, result_queue)
        except queue.Empty:
            break
        finally:
            proxy_queue.task_done()

# Ana fonksiyon
def main():
    start_time = time.time()
    
    # Proxy listesini oku
    proxies = read_proxies("utils/proxy_list.txt")
    print(f"Toplam {len(proxies)} proxy kontrol edilecek.")
    
    # Kuyrukları oluştur
    proxy_queue = queue.Queue()
    result_queue = queue.Queue()
    
    # Proxy'leri kuyruğa ekle
    for proxy in proxies:
        proxy_queue.put(proxy)
    
    # İş parçacıklarını oluştur ve başlat
    threads = []
    thread_count = min(MAX_THREADS, len(proxies))
    
    print(f"{thread_count} iş parçacığı başlatılıyor...")
    for _ in range(thread_count):
        thread = threading.Thread(target=worker, args=(proxy_queue, result_queue))
        thread.daemon = True
        thread.start()
        threads.append(thread)
    
    # Tüm iş parçacıklarının bitmesini bekle
    for thread in threads:
        thread.join()
    
    # Sonuçları topla
    valid_proxies = []
    while not result_queue.empty():
        proxy, _ = result_queue.get()
        valid_proxies.append(proxy)
    
    # Sonuçları dosyaya yaz
    with open(RESULT_FILE, 'w') as f:
        for proxy in valid_proxies:
            f.write(f"{proxy}\n")
    
    elapsed_time = time.time() - start_time
    print(f"\nTamamlandı! {len(valid_proxies)} geçerli proxy bulundu.")
    print(f"Geçerli proxy'ler '{RESULT_FILE}' dosyasına kaydedildi.")
    print(f"Toplam süre: {elapsed_time:.2f} saniye")

if __name__ == "__main__":
    main()