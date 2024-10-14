package excelimporter.reader.readers;

import com.mendix.core.Core;
import com.mendix.logging.ILogNode;
import com.mendix.replication.MendixReplicationException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSheet;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.function.Predicate;

public class ExcelXLSXDataReader {
    public static final ILogNode logNode = Core.getLogger("ExcelXLSXDataReader");
    private static final DataFormatter formatter = new DataFormatter();

    private ExcelXLSXDataReader() {
    }

    public static List<ExcelColumn> readHeaderRow(File excelFile, int sheetIndex, int headerRowIndex) throws ExcelImporterException {
        List<ExcelColumn> headerRow = new ArrayList<>();
        parseExcelFile(excelFile, sheetIndex, headerRowIndex, headerRow, 0, null, null);
        return headerRow;
    }

    public static long readDataRows(File excelFile, int sheetIndex, int startRowIndex, ExcelRowProcessor rowProcessor, Predicate<String> isColumnUsed) throws ExcelImporterException {
        parseExcelFile(excelFile, sheetIndex, 0, null, startRowIndex, rowProcessor, isColumnUsed);
        return rowProcessor.getRowCounter();
    }

    private static void parseExcelFile(File excelFile, int sheetIndex, int headerRowIndex, List<ExcelColumn> headerRow,
                                      int startRowIndex, ExcelRowProcessor rowProcessor, Predicate<String> isColumnUsed) throws ExcelImporterException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(excelFile) {
            @Override
            public void parseSheet(java.util.Map<String, XSSFSheet> shIdMap, CTSheet ctSheet) {
                // skipping parsing of any sheet
            }
        }) {
            try (var opcPackage = workbook.getPackage()) {
                var strings = new ReadOnlySharedStringsTable(opcPackage, false);
                ContentHandler handler;
                if (rowProcessor != null) {
                    handler = new ExtendedXSSFSheetXMLHandler(workbook.getStylesSource(), strings,
                            createSheetHandlerForData(sheetIndex, startRowIndex, rowProcessor, isColumnUsed),
                            formatter, false);
                } else {
                    handler = new ExtendedXSSFSheetXMLHandler(workbook.getStylesSource(), strings,
                            createSheetHandlerForHeader(headerRowIndex, headerRow), formatter, false);
                }
                XMLReader sheetParser = XMLHelper.newXMLReader();
                sheetParser.setContentHandler(handler);
                ArrayList<PackagePart> sheets = opcPackage.getPartsByContentType(XSSFRelation.WORKSHEET.getContentType());
                try (var sheet = sheets.get(sheetIndex).getInputStream()) {
                    InputSource sheetSource = new InputSource(sheet);
                    sheetParser.parse(sheetSource);
                }
            }
        }
        catch (XLSXHeaderFoundException e) {
            // safe to ignore this exception
        }
        catch (SAXException | ParserConfigurationException | IOException | InvalidFormatException e) {
            throw new ExcelImporterException("Error while opening workbook:" , e);
        } finally {
            if (rowProcessor != null) {
                handleRowProcessorCompletion(rowProcessor);
            }
        }
    }

    private static void handleRowProcessorCompletion(ExcelRowProcessor rowProcessor) {
        try {
            rowProcessor.finish();
            logRowProcessingResult(rowProcessor.getRowCounter());
        } catch (MendixReplicationException e) {
            throw new ExcelRuntimeException(e); // needed for backward compatibility
        }
    }

    private static void logRowProcessingResult(long rowCount) {
        if (rowCount == 0)
            logNode.warn("Excel Importer could not import any rows. Please check if the template is configured correctly. If the file was not created with Microsoft Excel for desktop, try opening the file with Excel and saving it with the same name before importing.");
        else
            logNode.info("Excel Importer successfully imported " + rowCount + " rows");
    }

    private static ExtendedXSSFSheetXMLHandler.SheetContentsHandler createSheetHandlerForHeader(int headerRowIndex, List<ExcelColumn> headerRow) {
        return new ExtendedXSSFSheetXMLHandler.SheetContentsHandler() {
            @Override
            public void startRow(int rowNum) {
                headerRow.clear();
            }

            @Override
            public void endRow(int rowNum) throws XLSXHeaderFoundException {
                if (rowNum == headerRowIndex && !headerRow.isEmpty() && !headerRow.stream().allMatch(Objects::isNull)) {
                    throw new XLSXHeaderFoundException("header row #" + (rowNum + 1));
                } else if (rowNum > headerRowIndex) {
                    throw new ExcelRuntimeException("Unable to find header row!!");
                }
            }

            @Override
            public void cell(String cellReference, String formattedValue, String rawValue, CellType cellType, String formatString, XSSFComment comment) {
                CellAddress cellAddr = new CellAddress(cellReference);
                if (isHeaderRow(cellAddr)) {
                    addToHeaderRow(cellAddr, cellType, formattedValue);
                }
            }

            private boolean isHeaderRow(CellAddress cellAddr) {
                return cellAddr.getRow() == headerRowIndex;
            }

            private void addToHeaderRow(CellAddress cellAddr, CellType cellType, String formattedValue) {
                int columnIndex = cellAddr.getColumn();
                switch (cellType) {
                    case FORMULA:
                    case STRING:
                        headerRow.add(new ExcelColumn(columnIndex, formattedValue));
                        break;
                    default:
                        headerRow.add(null);
                }
            }
        };
    }

    private static ExtendedXSSFSheetXMLHandler.SheetContentsHandler createSheetHandlerForData(int sheetIdx, int startRowIndex, ExcelRowProcessor rowProcessor, Predicate<String> isColumnUsed) {
        return new ExtendedXSSFSheetXMLHandler.SheetContentsHandler() {
            final ArrayList<ExcelRowProcessor.ExcelCellData> data = new ArrayList<>();
            boolean isNewRowStarted = false;

            @Override
            public void startRow(int rowNum) {
                isNewRowStarted = true;
                data.clear();
            }

            @Override
            public void endRow(int rowNum) throws SAXException {
                if (!data.isEmpty()) {
                    processRowData(rowNum);
                    data.clear();
                }
                isNewRowStarted = false;
            }

            private void processRowData(int rowNum) {
                try {
                    ExcelRowProcessor.ExcelCellData[] rowData = data.toArray(new ExcelRowProcessor.ExcelCellData[0]);
                    rowProcessor.processValues(rowData, rowNum, sheetIdx);
                } catch (MendixReplicationException e) {
                    throw new ExcelRuntimeException("Unable to process Excel row #" + (rowNum + 1) + " @Sheet #" + sheetIdx, e);
                }
            }

            @Override
            public void cell(String cellReference, String formattedValue, String rawValue, CellType cellType, String formatString, XSSFComment comment) {
                var cellAddr = new CellAddress(cellReference);
                int columnIndex = cellAddr.getColumn();
                try {
                    if (cellAddr.getRow() >= startRowIndex && isColumnUsed.test(String.valueOf(columnIndex))) {
                        if (logNode.isTraceEnabled())
                            logNode.trace(String.format("Reading %s / '%s' / %s", cellReference, rawValue, cellType));
                        if (rawValue == null) {
                            data.add(null);
                            return;
                        }
                        switch (cellType) {
                            case BOOLEAN:
                                data.add(new ExcelRowProcessor.ExcelCellData(columnIndex, rawValue, Integer.parseInt(rawValue) == 1));
                                break;
                            case ERROR:
                                // imported as null, because this can be handled in Mendix
                                if (rawValue.startsWith("#")) // Check if the error is due to a formula
                                {
                                    ExcelReader.logNode.error("Unable to import data due to invalid formula at cell address " + cellAddr);
                                    throw new ExcelRuntimeException("Unable to import data due to invalid formula at Excel row #" + (cellAddr.getRow() + 1));
                                }
                                data.add(new ExcelRowProcessor.ExcelCellData(columnIndex, rawValue, "ERROR:" + rawValue));
                                break;
                            case FORMULA:
                                data.add(new ExcelRowProcessor.ExcelCellData(columnIndex, rawValue, rawValue));
                                break;
                            case STRING: // We haven't seen this yet.
                                data.add(new ExcelRowProcessor.ExcelCellData(columnIndex, rawValue, formattedValue));
                                break;
                            case NUMERIC:
                                if (formatString != null) {
                                    final double dblCellValue = Double.parseDouble(rawValue);
                                    if (logNode.isTraceEnabled())
                                        logNode.trace(String.format("Formatting %s / '%s' using format: '%s' as %s", cellReference, rawValue, formatString, formattedValue));
                                    data.add(new ExcelRowProcessor.ExcelCellData(columnIndex, dblCellValue, formattedValue, formatString));
                                } else {
                                    data.add(new ExcelRowProcessor.ExcelCellData(columnIndex, rawValue, null));
                                }
                                break;
                            default:
                                data.add(null);
                        }
                    }
                } catch (Exception e) {
                    throw new ExcelRuntimeException(String.format("Unable to read Excel row #%d and cell #%d @Sheet #%d", cellAddr.getRow() + 1, columnIndex + 1, sheetIdx), e);
                }
            }
        };
    }
}
