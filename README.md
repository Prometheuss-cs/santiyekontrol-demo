# ŞantiyeKontrol — İnşaat Muhasebesi & Excel Otomasyon Demosu

İnşaat firmalarındaki muhasebe ekiplerinin her gün Excel'de saatlerini harcadığı üç temel işi tek tıkla çözen, **%100 yerel (internetsiz)** çalışan demo uygulamasıdır.

## Neler Var?

1. **İrsaliye – Fatura Çapraz Kontrolü**
   - Şantiye irsaliye Excel'i ile tedarikçi fatura Excel'ini saniyeler içinde karşılaştırır.
   - **Miktar farkı**, **birim fiyat artışı**, **mükerrer (çift) fatura**, **irsaliyesiz hayalet fatura** ve **fatura bekleyen irsaliyeleri** otomatik yakalar.
   - Sütun isimleri farklı olan Excel dosyaları için **otomatik sütun eşleştirme** desteği sunar.
   - Renkli & çok sayfalı `.xlsx` raporu ve PDF/yazdırılabilir yönetici özeti verir.

2. **Şantiye Maliyet & Tedarikçi Güvenilirlik Özeti**
   - Şantiye bazlı toplam maliyet ve yakalanan risk tutarlarını gösterir.
   - Malzeme gruplarına (Hazır Beton, Demir, Nakliye, Agrega vb.) göre dağılım çıkarır.
   - Hangi tedarikçinin daha sık hatalı/fazla fatura gönderdiğini listeler.

3. **Taşeron Hakediş & Kesinti Hesaplayıcı**
   - Yapılan iş kalemlerinden yola çıkarak **KDV, KDV Tevkifatı (4/10, 9/10 vb.), Teminat Kesintisi, Stopaj ve Avans Mahsubunu** hata payı olmadan hesaplar.
   - Tek tıkla icmal Excel (`.xlsx`) ve imzaya hazır PDF çıktısı üretir.

## Nasıl Açılır?

Hiçbir kurulum veya Python/Node paketi gerekmez:
1. `index.html` dosyasına **çift tıklayarak** tarayıcıda (Chrome / Edge) açın.
2. İlk ekranda **"Hazır Örnek Veriyle Demo Başlat"** butonuna basın; sistem hemen örnek verilerle 4 kritik hatayı yakalayacaktır.
3. Gerçek Excel sürükle-bırak testini göstermek için `ornek-dosyalar/` klasöründeki `santiye_irsaliyeleri.xlsx` ve `tedarikci_faturalari.xlsx` dosyalarını kullanabilirsiniz.

## Gizlilik

Tüm Excel ve CSV dosyaları **yalnızca kullanıcının kendi bilgisayarında / tarayıcısında** işlenir; hiçbir veri internete veya harici sunuculara gönderilmez.
