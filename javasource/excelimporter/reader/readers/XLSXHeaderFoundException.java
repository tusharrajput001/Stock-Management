package excelimporter.reader.readers;

import org.xml.sax.SAXException;

public class XLSXHeaderFoundException  extends SAXException {
    XLSXHeaderFoundException (String message) {
        super(message);
    }
}
