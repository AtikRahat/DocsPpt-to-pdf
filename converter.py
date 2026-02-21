import os
import sys
import comtypes.client
from win32com.client import constants, gencache

def convert_docx_to_pdf(input_path, output_path, start_page=None, end_page=None):
    """
    Converts a DOCX file to PDF using Microsoft Word.
    """
    word = None
    doc = None
    try:
        word = gencache.EnsureDispatch('Word.Application')
        word.Visible = False
    except Exception:
        import win32com.client
        word = win32com.client.Dispatch('Word.Application')
        word.Visible = False

    try:
        input_path = os.path.abspath(input_path)
        output_path = os.path.abspath(output_path)

        doc = word.Documents.Open(input_path, ReadOnly=1)
        
        wdExportFormatPDF = 17 
        wdExportAllDocument = 0
        wdExportFromTo = 3
        
        export_range = wdExportAllDocument
        if start_page is not None and end_page is not None:
            export_range = wdExportFromTo
        
        if start_page: start_page = int(start_page)
        if end_page: end_page = int(end_page)

        doc.ExportAsFixedFormat(
            OutputFileName=output_path,
            ExportFormat=wdExportFormatPDF,
            OpenAfterExport=False,
            OptimizeFor=0,
            Range=export_range,
            From=start_page if start_page else 0,
            To=end_page if end_page else 0,
            Item=0,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=0,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False
        )
        return True, "Conversion successful"
    except Exception as e:
        return False, str(e)
    finally:
        if doc:
            doc.Close(SaveChanges=0)
        if word:
            word.Quit()

def convert_pptx_to_pdf(input_path, output_path, start_page=None, end_page=None):
    """
    Converts a PPTX/PPT file to PDF using Microsoft PowerPoint.
    """
    powerpoint = None
    presentation = None
    try:
        powerpoint = gencache.EnsureDispatch('PowerPoint.Application')
    except Exception:
        import win32com.client
        powerpoint = win32com.client.Dispatch('PowerPoint.Application')

    try:
        input_path = os.path.abspath(input_path)
        output_path = os.path.abspath(output_path)

        presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)
        
        # https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppsaveasfiletype
        ppSaveAsPDF = 32
        
        # https://learn.microsoft.com/en-us/office/vba/api/powerpoint.ppprintrange
        # Note: PowerPoint ExportAsFixedFormat is slightly different from Word.
        # It's often easier to use SaveAs for simple cases, 
        # but ExportAsFixedFormat allows more control.
        # https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.exportasfixedformat
        
        # ppFixedFormatTypePDF = 2
        # ppFixedFormatIntentScreen = 1
        
        if start_page is not None and end_page is not None:
            # PowerPoint page range is slightly more involved via PrintOptions or ExportAsFixedFormat
            # For simplicity in this tool, we use SaveAs if no range, 
            # or try ExportAsFixedFormat for range.
            presentation.ExportAsFixedFormat(
                Path=output_path,
                FixedFormatType=2, # ppFixedFormatTypePDF
                RangeType=4, # ppPrintFromTo
                From=int(start_page),
                To=int(end_page)
            )
        else:
            presentation.SaveAs(output_path, ppSaveAsPDF)
            
        return True, "Conversion successful"
    except Exception as e:
        return False, str(e)
    finally:
        if presentation:
            presentation.Close()
        if powerpoint:
            powerpoint.Quit()

def convert_to_pdf(input_path, output_path, start_page=None, end_page=None):
    """
    Dispatcher based on file extension.
    """
    ext = os.path.splitext(input_path)[1].lower()
    if ext in ['.docx', '.doc']:
        return convert_docx_to_pdf(input_path, output_path, start_page, end_page)
    elif ext in ['.pptx', '.ppt']:
        return convert_pptx_to_pdf(input_path, output_path, start_page, end_page)
    else:
        return False, f"Unsupported file format: {ext}"
