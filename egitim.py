import os
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

import torch
from ultralytics import YOLO

def main():
    # GPU (CUDA) kontrolü yapıyoruz
    cuda_available = torch.cuda.is_available()
    device_to_use = 0 if cuda_available else 'cpu'
    if not cuda_available:
        print("\n[UYARI] NVIDIA GPU (CUDA) tespit edilemedi! Eğitim CPU üzerinde çalışacak ve oldukça YAVAŞ sürecektir.")
        print("GPU kullanabilmek için CUDA destekli PyTorch yüklemeniz önerilir:")
        print("pip install torch torchvision torchaudio --index-url https://download.pytorch.org/whl/cu121 --force-reinstall\n")

    # --- EN SON EĞİTİLEN AĞIRLIKLARI ARAMA (GEREKSİZ EĞİTİMİ ÖNLEMEK İÇİN) ---
    latest_best_pt = None
    yolo_egitim_dir = os.path.join('runs', 'detect', 'YOLO_Egitim')
    if os.path.exists(yolo_egitim_dir):
        import re
        subdirs = [d for d in os.listdir(yolo_egitim_dir) if d.startswith('Sentetik_Model')]
        if subdirs:
            # Doğru sıralama için nümerik sıralama (Sentetik_Model, Sentetik_Model-2, ...)
            subdirs.sort(key=lambda x: [int(c) if c.isdigit() else c for c in re.split(r'(\d+)', x)])
            # En son güncel olandan geriye doğru weights/best.pt arıyoruz
            for d in reversed(subdirs):
                pt_path = os.path.join(yolo_egitim_dir, d, 'weights', 'best.pt')
                if os.path.exists(pt_path):
                    latest_best_pt = pt_path
                    break

    # --- AŞAMA 1: SENTETİK VERİ İLE ÖN EĞİTİM ---
    if latest_best_pt:
        print(f"\n[BİLGİ] Daha önce başarıyla eğitilmiş ön eğitim ağırlık dosyası bulundu: {latest_best_pt}")
        print("1. Aşama (Sentetik veri ön eğitimi) atlanıyor, doğrudan 2. Aşamaya geçiliyor...\n")
        best_weights_path = latest_best_pt
    else:
        print("1. Aşama: Sentetik verilerle ön eğitim başlıyor...")
        # YOLOv8'in en hafif ve hızlı modeli olan Nano'yu (n) yüklüyoruz
        base_model = YOLO('yolov8n.pt') 
        
        # Modeli sentetik verilerle eğitiyoruz
        # 6GB VRAM kapasitesini taşırmamak için batch boyutunu 16 olarak ayarlıyoruz. 
        # Eğer bellek hatası alırsan burayı 8 yapabilirsin.
        results = base_model.train(
            data='data_sentetik.yaml',
            epochs=50,       
            imgsz=640,       
            batch=16,        
            device=device_to_use,  # GPU varsa 0 (GPU), yoksa 'cpu' kullanır
            project='YOLO_Egitim',
            name='Sentetik_Model'
        )
        best_weights_path = os.path.join(results.save_dir, 'weights', 'best.pt')

    # --- AŞAMA 2: GERÇEK VERİ İLE FINE-TUNING (İNCE AYAR) ---
    print("\n2. Aşama: Gerçek verilerle Fine-Tuning başlıyor...")
    
    # Eğittiğimiz ve en iyi sonucu veren ağırlık dosyasını (best.pt) yüklüyoruz
    print(f"En iyi ağırlıklar yükleniyor: {best_weights_path}")
    fine_tune_model = YOLO(best_weights_path)
    
    # Bu modeli gerçek verilerle ince ayardan geçiriyoruz
    fine_tune_model.train(
        data='data_gercek.yaml',
        epochs=100,      # Gerçek veriyi iyice kavraması için epok sayısını 100 tutuyoruz
        imgsz=640,
        batch=16,
        device=device_to_use,
        project='YOLO_Egitim',
        name='Final_Hibrit_Model'
    )
    
    print("\nEğitim tamamlandı! Final modelin şurada: YOLO_Egitim/Final_Hibrit_Model/weights/best.pt")

if __name__ == '__main__':
    # Windows ortamında multiprocessing (çoklu işlem) hatalarını önlemek için main bloğu şarttır
    main()