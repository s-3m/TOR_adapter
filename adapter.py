import pandas as pdgit
import json


def city_adapter():
    data = pd.read_excel("Справочник НП.xlsx").to_dict("records")
    for item in data[:6]:
        result = {"meta": {"metaInfo": {}, "xsd": None}, "items": []}
        result["items"].append({"code": item["Код населенного пункта"]})
        result["items"].append({"city_name": item["Наименование населенного пункта"]})
        result["items"].append({"parentEntries": "nsiUnifiedSPR_CITIES_EO_AGREGATOR"})
        result["items"].append({"dictionaryType": "unified"})
        result["items"].append({"dictionaryUnitId": None})
        result["items"].append({"region_code": {"code": item["Код населенного пункта"].split(".")[0], "name": item["Наименование региона"]}})
        result["items"].append({"department_code": {"code": item["Уникальный код ведомства, ответственного за оказание муниципальных услуг в ритуальной сфере"], "name": item["Наименование ведомства, ответственного за оказание муниципальных услуг в ритуальной сфере"]}})

        with open(f"{item["Код населенного пункта"]}.json", "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=4)


def dep_adapter():
    data = pd.read_excel("Справочник ведомств.xlsx").to_dict("records")
    count = 1
    xls_finish_result = []
    for item in data[:1]:
        result = {"items": []}
        xls_result = {}

        result["items"].append({"code": item["Уникальный код ведомства"]})
        xls_result["code"] = item["Уникальный код ведомства"]
        xls_result["name"] = item["Наименование ведомства"]
        result["items"].append({"name": item["Наименование ведомства"]})

        result["items"].append({"dictionaryType": "unified"})
        xls_result["dictionaryType"] = "unified"
        result["items"].append({"dictionaryUnitId": None})
        xls_result["dictionaryUnitId"] = None

        dep_info = {
            "department_info": {
                "e_mail": item["Адрес электронной почты ведомства"],
                "region": {"code": item["Уникальный код ведомства"].split(".")[0], "name": item["Наименование региона"]},
                "telephone": item["Телефон ведомства"],
                "manager_fio": item["ФИО руководителя ведомства"].title().strip(),
                "post_address": {
                    "okato": "",
                    "oktmo": "",
                    "ifnsfl": "",
                    "ifnsul": "",
                    "country": "",
                    "postalCode": item["Почтовый индекс ведомства"],
                    "regionCode": "",
                    "fullAddress": f"{item["Почтовый индекс ведомства"]}, {item["Почтовый адрес ведомства"]}",
                    "addressParts": [],
                    "addressAsObject": {},
                    "cadastralNumber": "",
                    "isCustomAddress": False,
                    "isSpecialAddress": False,
                    "unrecognizablePart": f"{item["Почтовый адрес ведомства"]}",
                },
                "department_name": item["Наименование ведомства"],
                "department_ogrn": item["ОГРН ведомства"],
                "manager_position": item["Должность руководителя"],
                "department_address": {
                    "okato": "",
                    "oktmo": "",
                    "ifnsfl": "",
                    "ifnsul": "",
                    "country": "",
                    "postalCode": item["Почтовый индекс ведомства"],
                    "regionCode": "",
                    "fullAddress": f"{item["Почтовый индекс ведомства"]}, {item["Почтовый адрес ведомства"]}",
                    "addressParts": [],
                    "addressAsObject": {},
                    "cadastralNumber": "",
                    "isCustomAddress": False,
                    "isSpecialAddress": False,
                    "unrecognizablePart": f"{item["Почтовый адрес ведомства"]}",
                },
            }
        }
        result["items"].append(dep_info)
        xls_result["department_info"] = f"{dep_info["department_info"]}"

        bur_info = {
            "department_burials": {
                "confession": {
                    "confessional_judaic": (
                        False
                        if item["Наличие иудейских захоронений"] == "Нет"
                        else True
                    ),
                    "confessional_islamic": (
                        False
                        if item["Наличие исламских захоронений"] == "Нет"
                        else True
                    ),
                    "confessional_buddhist": (
                        False
                        if item["Наличие буддийских захоронений"] == "Нет"
                        else True
                    ),
                    "confessional_catholic": (
                        False
                        if item["Наличие католических захоронений"] == "Нет"
                        else True
                    ),
                    "confessional_orthodox": (
                        False
                        if item["Наличие православных захоронений"] == "Нет"
                        else True
                    ),
                    "confessional_protestant": (
                        False
                        if item["Наличие протестанских захоронений"] == "Нет"
                        else True
                    ),
                },
                "specialization": {
                    "specialized_honorary": (
                        False if item["Наличие почётных захоронений"] == "Нет" else True
                    ),
                    "specialized_military": (
                        False if item["Наличие воинских захоронений"] == "Нет" else True
                    ),
                    "specialized_walk_of_fame": (
                        False if item["Наличие Аллеи Славы"] == "Нет" else True
                    ),
                },
                "walls_of_sorrow": (
                    False if item["Наличие стен скорби (колумбария)"] == "Нет" else True
                ),
                "specialized_burials": (
                    True
                    if item["Наличие почётных захоронений"] == "Да"
                    or item["Наличие воинских захоронений"] == "Да"
                    or item["Наличие Аллеи Славы"] == "Да"
                    else False
                ),
                "burial_choice_access": (
                    False if item["Доступ к выбору мест захоронения"] == "Нет" else True
                ),
                "confessional_burials": True,
            }
        }
        result["items"].append(bur_info)
        xls_result["department_burials"] = f"{bur_info["department_burials"]}"

        xls_finish_result.append(xls_result)

        with open(f"{item["Уникальный код ведомства"]}.json", "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=4)

        count += 1

    pd.DataFrame(xls_finish_result).to_excel(f"REG999.DEP{count}.xlsx", index=False)
    pd.DataFrame(xls_finish_result).to_csv(f"REG999.DEP{count}.csv", index=False)


def search_dot():
    count = 0
    data = pd.read_excel("dep3.xlsx").to_dict("records")

    for item in data:
        name = item["code"]
        result = {"items": []}
        try:
            department_info = json.loads(item["department_info"])
            department_burials = json.loads(item["department_burials"])
            if "." in department_info["department_ogrn"]:
                department_ogrn = (
                    department_info.get("department_ogrn").replace(".", "").strip()
                )
                department_info["department_ogrn"] = department_ogrn
                item["department_info"] = department_info
                item["department_burials"] = department_burials
                for key in item:
                    if key in (
                        "_id",
                        "auid",
                        "guid",
                        "code",
                        "name",
                        "dictionaryType",
                        "department_info",
                        "dictionaryUnitId",
                        "department_burials",
                    ):
                        if key == "dictionaryUnitId":
                            result["items"].append({key: None})
                        else:
                            result["items"].append({key: item[key]})

                with open(f"dep/{name}.json", "w", encoding="utf-8") as f:
                    json.dump(result, f, ensure_ascii=False, indent=4)
                count += 1
        except json.decoder.JSONDecodeError:
            print(name)
    print(count)

    # pd.DataFrame(data).to_excel(f"dep_clean.xlsx", index=False)


if __name__ == "__main__":
    dep_adapter()
