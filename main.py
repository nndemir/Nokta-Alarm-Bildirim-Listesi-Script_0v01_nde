import pandas as pd
import ijson
import logging
import os

logging.basicConfig(level=logging.INFO)


def load_event_handlers(eventHandlersJSON):
    mapping = {}
    with open(eventHandlersJSON, "r", encoding="utf-8") as f:
        for handler in ijson.items(f, "eventHandlers.item"):
            name = handler.get("name", "")
            for et in handler.get("eventTypes", []):
                xid = et.get("detectorXID")
                if xid:
                    mapping.setdefault(xid, []).append(name)
    return mapping


def dataPointsJsonReaderFunc(dataPointsJSON, eventHandlersJSON):
    data_points = []
    handlers_map = load_event_handlers(eventHandlersJSON)
    count = 0
    with open(dataPointsJSON, "r", encoding="utf-8") as f:
        for item in ijson.items(f, "dataPoints.item"):
            eventDetectors = item.get("eventDetectors", [])
            tags = item.get("tags", {})
            # print(tags)
            # if count > 1000:
            #     break
            if eventDetectors:
                try:
                    for ed in eventDetectors:
                        ED_Zamani = ""
                        state_text = ""

                        match ed.get("type"):
                            case "MULTISTATE_STATE":
                                state = ed.get("state")
                                for tr in item.get("textRenderer", {}).get(
                                    "multistateValues", []
                                ):
                                    if tr.get("key") == state:
                                        state_text = tr.get("text", "")
                            case "BINARY_STATE":
                                state = ed.get("state")
                                if ed.get("state"):
                                    state_text = state_text = item.get(
                                        "textRenderer", {}
                                    ).get("oneLabel", "")
                                else:
                                    state_text = state_text = item.get(
                                        "textRenderer", {}
                                    ).get("zeroLabel", "")

                        ED_Zamani = (
                            f"{ed.get('duration')} {ed.get('durationType') or ''}"
                        )

                        data_points.append(
                            {
                                "xid": item.get("xid"),
                                "deviceName": item.get("deviceName"),
                                "name": item.get("name"),
                                "Veri_Tipi": item.get("pointLocator", {}).get(
                                    "dataType", ""
                                ),
                                # "Kampus": tags.get("Kampus", ""),
                                # "Modul": tags.get("Modul", ""),
                                # "Cihaz_Tipi": tags.get("Cihaz Tipi", ""),
                                # "CI_Kodu": tags.get("CI Kodu", ""),
                                # "Bildirim Tagi": tags.get("Bildirim", ""),
                                **tags,
                                "ED_xid": ed.get("xid", ""),
                                "ED_Type": ed.get("type", ""),
                                "ED_Level": ed.get("alarmLevel", ""),
                                "ED_Zamani": ED_Zamani,
                                "ED_Limit": ed.get("limit", ""),
                                "ED_Reset_Limit": ed.get("resetLimit", ""),
                                "ED_States": state_text,
                                "ED_En_Dusuk_Seviye": ed.get("low", ""),
                                "ED_En_Yuksek_Seviye": ed.get("high", ""),
                                "Secili Bildirimler": ",".join(
                                    handlers_map.get(ed.get("xid"), [])
                                ),
                            }
                        )
                except Exception as e:
                    logging.error(
                        f"Hata: {e} | Satir: {count} | XID: {item.get('xid')} | Cihaz Tipi: {item.get('deviceName')} | Name: {item.get('name')}"
                    )
            else:  # hiç eventDetectors yoksa
                data_points.append(
                    {
                        "xid": item.get("xid"),
                        "deviceName": item.get("deviceName"),
                        "name": item.get("name"),
                        "Veri_Tipi": item.get("pointLocator", {}).get("dataType", ""),
                        # "Kampus": tags.get("Kampus", ""),
                        # "Modul": tags.get("Modul", ""),
                        # "Cihaz_Tipi": tags.get("Cihaz Tipi", ""),
                        # "CI_Kodu": tags.get("CI Kodu", ""),
                        # "Bildirim Tagi": tags.get("Bildirim", ""),
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


import os
import pandas as pd


def main():
    dataPointsJSONPath = "./data-points.json"
    eventHandlersJSONPath = "./event-handlers.json"

    # Çıktı klasörü
    output_dir = "./excel_outputs"
    os.makedirs(output_dir, exist_ok=True)  # klasör yoksa oluştur

    # Temel dosya adı
    base_name = "MI-Nokta-Alarm-Bildirim_Listesi"
    version = "0v02"
    fileNum = 1

    # İlk dosya yolu (numara yok)
    file_path = os.path.join(output_dir, f"{base_name}_{version}_nde.xlsx")

    # Eğer dosya varsa numara ekle
    while os.path.exists(file_path):
        fileNum += 1
        file_path = os.path.join(
            output_dir, f"{base_name}_{version}_nde-{fileNum}.xlsx"
        )

    # DataFrame oluştur
    data_points = dataPointsJsonReaderFunc(dataPointsJSONPath, eventHandlersJSONPath)
    df = pd.DataFrame(data_points)

    # Excel olarak kaydet
    print("Excel dosyası oluşturuluyor, lütfen bekleyin...")
    df.to_excel(file_path, index=False, engine="openpyxl")
    print("Dosya oluşturuldu:", file_path)


if __name__ == "__main__":
    main()
