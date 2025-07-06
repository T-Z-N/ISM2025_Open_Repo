import pandas as pd

# CSV dosyasının adı
csv_file_path = 'query-result.csv'
# Oluşturulacak Excel dosyasının adı
excel_file_path = 'class_annotations.xlsx'

try:
    # CSV dosyasını oku
    df = pd.read_csv(csv_file_path)

    # NaN değerleri (boş hücreler) boş string ile değiştir (görsel olarak daha iyi)
    df.fillna('', inplace=True)

    # Excel yazıcısı oluştur
    with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Annotations', index=False)

        # Çalışma sayfasını al
        worksheet = writer.sheets['Annotations']

        # Hücreleri birleştirmek için mantık
        # İlk satır başlık olduğu için veri satırları 2'den başlar (1-indexed)
        # DataFrame 0-indexed olduğu için buna dikkat etmeliyiz.

        if not df.empty:
            current_class_uri = None
            merge_start_row = 2  # Excel'de veri için ilk satır (başlıktan sonra)

            for idx in range(len(df)):
                # Excel'deki geçerli satır numarası (1-tabanlı)
                # idx (0-tabanlı DataFrame indeksi) + 2 (1 başlık satırı + 1 taban farkı)
                excel_current_row = idx + 2

                # Eğer bu ilk satırsa veya mevcut sınıf bir öncekiyle aynı değilse
                if current_class_uri is None or df.loc[idx, 'class'] != current_class_uri:
                    # Bir önceki sınıf bloğu için birleştirme yap (eğer birden fazla satır kaplıyorsa)
                    if current_class_uri is not None and excel_current_row - 1 > merge_start_row:
                        worksheet.merge_cells(start_row=merge_start_row, start_column=1,
                                              end_row=excel_current_row - 1, end_column=1)

                    # Yeni sınıf bloğu için başlangıcı ayarla
                    current_class_uri = df.loc[idx, 'class']
                    merge_start_row = excel_current_row

            # Döngüden sonra son bloğu da birleştir
            if current_class_uri is not None and (len(df) + 1) > merge_start_row: # len(df)+1 son satırın bir altı
                worksheet.merge_cells(start_row=merge_start_row, start_column=1,
                                      end_row=len(df) + 1, end_column=1)

        print(f"Excel dosyası '{excel_file_path}' başarıyla oluşturuldu.")

except FileNotFoundError:
    print(f"Hata: '{csv_file_path}' dosyası bulunamadı.")
except Exception as e:
    print(f"Bir hata oluştu: {e}")