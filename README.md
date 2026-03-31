# Any2PDF Merger 📄✨

**Any2PDF Merger**, bilgisayarınızdaki farklı türdeki sayısız dosyayı **hiçbir kalite veya görsel kaybı olmadan** tam sayfa olarak saniyeler içerisinde PDF’e dönüştüren ve tek bir dosya altında "sıralı olarak" birleştiren harika bir Python masaüstü uygulamasıdır.

Özellikle yazılımcılar, araştırmacılar, makine öğrenmesi/derin öğrenme öğrencileri ve farklı formatlardaki ders notlarını tek bir PDF'te toplamak isteyenler için birebirdir.

## Özellikler

- **Jupyter Notebook (`.ipynb`) Tam Destek**: Çıktılar, grafikler, neural network loss haritaları, markdown formülleri ve output hücrelerini hiçbir detayı atlamadan orijinal web kalitesinde renderlar.
- **Web Sayfası Desteği (`.html`, `.htm`)**: Görünmez bir Microsoft Edge motoru arka planda HTML dosyalarınızı açar ve kusursuz baskı kalitesinde PDF formatına işler. Kalite bozulması veya kesilme yaşanmaz!
- **Görseller İçin %100 Piksel Uyumu**: Orijinal `PNG, JPG, BMP` vb. görseller kırpılmadan ve hiçbir sıkıştırma (compression) kaybı yaşanmadan direkt kendi orijinal piksel boyutunda içe aktarılır.
- **Tüm Office/Metin Dosyaları Desteklenir**: `DOCX, PPTX, XLSX`, Python/C++ kod dosyalarınızı tek tıkla topluca sürükleyip PDF'in içine yerleştirebilirsiniz.
- **Sıralama Arayüzü**: Sürükleyip bıraktıktan sonra uygulama içindeki oklar sayesinde sayfaların çıktı sırasını dilediğiniz gibi ayarlayabilirsiniz. 
- **Modern ve Sade Arayüz**: Dark-mode ağırlıklı, kullanımı çocuk oyuncağı olan oldukça şık arayüz (Tkinter/Windows 11 API).

## 🛠️ Kurulum ve Ön Gereksinimler

Projenin bağımlılıklarını kurmak için bir terminal açın ve aşağıdaki repoyu bilgisayarınıza indirip klasöre girerek şu komutu çalıştırın:

```cmd
pip install Pillow pypdf reportlab nbformat nbconvert comtypes
```

*Not: HTML ve `.ipynb` özelliklerinin kusursuz çalışabilmesi için işletim sisteminizde Microsoft Edge'in yüklü olması yeterlidir (Chrome tabanlı). Office dosyaları (`.docx`, `.pptx`) işlemleri için ise sistemde Microsoft Office'in yüklü olması gerekir.*

## Nasıl Kullanılır?

Sadece aşağıdaki komutu çalıştırarak uygulamayı açmanız yeterlidir:

```cmd
python pdf_birlestirici.py
```
Açılan pencerede:
1. `+ Dosya Ekle` butonuyla PDF'e dahil edilecek evrak/görsel/kod/jupyter notlarınızı seçin.
2. Sırasını değiştirmek istediklerinizi yan oklarla organize edin.
3. ` PDF Olarak Birleştir` butonuna basın! Sadece birkaç saniye sürecek.

## 🤝 Katkıda Bulunma
Her türlü öneri, hata çözümü (PR) ve geliştirmeye tamamen açıktır! Repoyu "Fork" yapıp istediğiniz gibi geliştirebilirsiniz.

