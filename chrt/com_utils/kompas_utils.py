import os
import re
import win32com.client 
from win32com.client import gencache
from utils.error_utils import handle_error, log_info
import pythoncom

kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)

def initialize_kompas():
    try:
        kompas_app = win32com.client.Dispatch("Kompas.Application.7")
        kompas_app.Visible = False
        return kompas_app
    except Exception as e:
        raise RuntimeError(f"Ошибка инициализации KOMPAS: {e}")

def save_to_version(kompas_app, file_path, output_path, version):
    try:
        doc = kompas_app.Documents.Open(file_path)
        success = kompas_app.ksSaveDocumentEx(output_path, version)
        doc.Close()
        return success
    except Exception as e:
        raise RuntimeError(f"Ошибка сохранения: {e}")

def export_to_pdf(kompas_app, file_path, output_path):
    try:
        doc = kompas_app.Documents.Open(file_path)
        converter = kompas_app.Converter("Pdf2d.dll")
        converter.Convert(file_path, output_path, 0, False)
        doc.Close()
        return True
    except Exception as e:
        raise RuntimeError(f"Ошибка экспорта в PDF: {e}")

def export_to_step(kompas_app, file_path, output_path, step_format="AP203"):
    try:
        doc = kompas_app.Documents.Open(file_path)
        doc.ExportToSTEP(output_path, step_format)
        return True
    except Exception as e:
        raise RuntimeError(f"Ошибка экспорта в STEP: {e}")
    finally:
        if 'doc' in locals():
            try:
                doc.Close()
            except:
                pass

def amount_sheet(doc7):
    sheets = {"A0": 0, "A1": 0, "A2": 0, "A3": 0, "A4": 0, "A5": 0}
    for sheet in range(doc7.LayoutSheets.Count):
        format = doc7.LayoutSheets.Item(sheet).Format  # sheet - номер листа, отсчёт начинается от 0
        sheets["A" + str(format.Format)] += 1 * format.FormatMultiplicity
    return sheets


# Прочитаем основную надпись чертежа
def stamp(doc7):
    for sheet in range(doc7.LayoutSheets.Count):
        style_filename = os.path.basename(doc7.LayoutSheets.Item(sheet).LayoutLibraryFileName)
        style_number = int(doc7.LayoutSheets.Item(sheet).LayoutStyleNumber)

        if style_filename in ['graphic.lyt', 'Graphic.lyt'] and style_number == 1:
            stamp = doc7.LayoutSheets.Item(sheet).Stamp
            return {"Scale": re.findall(r"\d+:\d+", stamp.Text(6).Str)[0],
                    "Designer": stamp.Text(110).Str}

    return 'Неопределенный стиль оформления'


# Подсчет технических требований, в том случае, если включена автоматическая нумерация
def count_demand(doc7, module7):
    IDrawingDocument = doc7._oleobj_.QueryInterface(module7.NamesToIIDMap['IDrawingDocument'], pythoncom.IID_IDispatch)
    drawing_doc = module7.IDrawingDocument(IDrawingDocument)
    text_demand = drawing_doc.TechnicalDemand.Text

    count = 0  # Количество пунктов технических требований
    for i in range(text_demand.Count):  # Прохоим по каждой строчке технических требований
        if text_demand.TextLines[i].Numbering == 1:  # и проверяем, есть ли у строки нумерация
            count += 1

    # Если нет нумерации, но есть текст
    if not count and text_demand.TextLines[0]:
        count += 1

    return count


# Подсчёт размеров на чертеже, для каждого вида по отдельности
def count_dimension(doc7, module7):
    IKompasDocument2D = doc7._oleobj_.QueryInterface(module7.NamesToIIDMap['IKompasDocument2D'],
                                                     pythoncom.IID_IDispatch)
    doc2D = module7.IKompasDocument2D(IKompasDocument2D)
    views = doc2D.ViewsAndLayersManager.Views

    count_dim = 0
    for i in range(views.Count):
        ISymbols2DContainer = views.View(i)._oleobj_.QueryInterface(module7.NamesToIIDMap['ISymbols2DContainer'],
                                                                    pythoncom.IID_IDispatch)
        dimensions = module7.ISymbols2DContainer(ISymbols2DContainer)

        # Складываем все необходимые раpмеры
        count_dim += dimensions.AngleDimensions.Count + \
                     dimensions.ArcDimensions.Count + \
                     dimensions.Bases.Count + \
                     dimensions.BreakLineDimensions.Count + \
                     dimensions.BreakRadialDimensions.Count + \
                     dimensions.DiametralDimensions.Count + \
                     dimensions.Leaders.Count + \
                     dimensions.LineDimensions.Count + \
                     dimensions.RadialDimensions.Count + \
                     dimensions.RemoteElements.Count + \
                     dimensions.Roughs.Count + \
                     dimensions.Tolerances.Count

    return count_dim

def analyze_drawings(kompas_app, file_path):
    try:
        doc = kompas_app.Documents.Open(file_path)
        result = {
            "filename": os.path.basename(file_path),
            "format": "Неизвестно",
            "scale": 1,
            "dimensions": 0,
            "technical_requirements": 0,
            "views": 0,
            "labor_intensity": 0,
        }

        # Проверка на 2D-документ
        if hasattr(doc, "DimensionCollection"):
            result["dimensions"] = doc.DimensionCollection.Count
            result["technical_requirements"] = doc.TechnicalRequirements.Count
            result["views"] = doc.ViewsAndLayersManager.Views.Count

            try:
                sheet = doc.LayoutSheets.ActiveSheet
                result["format"] = getattr(sheet, "FormatName", "Неизвестно")
                result["scale"] = getattr(sheet, "Scale", 1)
            except Exception as e:
                log_info(f"Ошибка при чтении параметров листа: {e}")

        else:
            log_info(f"Документ {file_path} не является 2D-чертёжом.")

        # Расчёт трудоёмкости
        result["labor_intensity"] = (
            result["dimensions"] * 0.5 +
            result["technical_requirements"] * 2 +
            result["views"] * 1.5
        )

        doc.Close()
        return result

    except Exception as e:
        raise RuntimeError(f"Ошибка при анализе чертежа: {e}")
