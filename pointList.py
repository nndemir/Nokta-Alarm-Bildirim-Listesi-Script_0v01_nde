#
#
#
# Project: TRC MI
# Task: All Data Points
#
#
#

import pandas as pd
import ijson
import logging
import os

logging.basicConfig(level=logging.INFO)


def dataPointsJsonReaderFunc(dataPointsJSON):
    data_points = []
    count = 0
    with open(dataPointsJSON, "r", encoding="utf-8") as f:
        for item in ijson.items(f, "dataPoints.item"):
            tags = item.get("tags", {})

            # if count > 200:
            #     break

            data_points.append(
                {
                    "xid": item.get("xid"),
                    "deviceName": item.get("deviceName"),
                    "name": item.get("name"),
                    "Veri_Tipi": item.get("pointLocator", {}).get("dataType", ""),
                    **tags,
                    "ED_xid": "",
                    "ED_Type": "",
                    "ED_Level": "",
                    "ED_Zamani": "",
                    "ED_Limit": "",
                    "ED_Reset_Limit": "",
                    "ED_States": "",
                    "ED_En_Dusuk_Seviye": "",
                    "ED_En_Yuksek_Seviye": "",
                    "Secili Bildirimler": "",
                }
            )

            if count % 1000 == 0:  # her 1000 kayıtta log yaz
                logging.info(f"{count} kayit işlendi")
            count += 1
        logging.info(f"{count} kayit işlendi")
    return data_points


def main():
    dataPointsJSONPath = "./data-points.json"
    output_dir = "./excel_outputs"
    os.makedirs(output_dir, exist_ok=True)

    base_name = "MI-Nokta_Listesi"
    fileNum = 1
    file_path = os.path.join(output_dir, f"{base_name}_nde.xlsx")
    while os.path.exists(file_path):
        fileNum += 1
        file_path = os.path.join(output_dir, f"{base_name}_nde-{fileNum}.xlsx")

    # DataFrame oluştur
    data_points = dataPointsJsonReaderFunc(dataPointsJSONPath)
    df = pd.DataFrame(data_points)

    # Boş Kampus değerlerini doldur
    df["Kampus"] = df["Kampus"].fillna("Bos")

    # Sadece belirli kampüsleri seçmek için liste
    lokasyonlar = [
        "Ankara DC",
        "Avrupa DC",
        "Kartal DC",
        "Manisa NDC2",
        "MASLAK42 NDC2",
        "Yenibosna NDC2",
        "MTK NDC2",
        "Adana NDC",
        "Başkent NDC",
        "Diyarbakır NDC",
        "Erzurum NDC",
        "Hatay NDC",
        "Kayseri NDC",
        "Malatya NDC",
        "Mugla NDC",
        "Trabzon NDC",
    ]

    # True ise sadece bu lokasyonları al
    nde_lokasyonlari = True
    if nde_lokasyonlari:
        # Büyük/küçük harf ve Türkçe karakter uyumunu sağlamak için
        df_filtered = df[
            df["Kampus"].str.strip().str.upper().isin([k.upper() for k in lokasyonlar])
        ]
    else:
        df_filtered = df

    print("Excel dosyası oluşturuluyor, lütfen bekleyin...")

    with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
        # Seçilen kampüsleri kendi sheet'lerine yaz
        for kampus, grup in df_filtered.groupby("Kampus"):
            grup.to_excel(writer, sheet_name=str(kampus)[:31], index=False)

        # Boş olanları ayrı sheet
        boslar = df[df["Kampus"] == "Bos"]
        if not boslar.empty:
            boslar.to_excel(writer, sheet_name="Bos", index=False)

    print("Dosya oluşturuldu:", file_path)


if __name__ == "__main__":
    main()
