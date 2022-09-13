
import net.sourceforge.tess4j.Tesseract;
import net.sourceforge.tess4j.TesseractException;
import org.apache.poi.hslf.record.CString;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.ooxml.POIXMLDocument;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.*;



import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.*;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class TextToDoc {
    public static void main(String[] args) throws IOException, IllegalArgumentException, TesseractException {

        File image = new File("C:\\Java_tesseract\\pass.jpg");
        Tesseract tesseract = new Tesseract();
        tesseract.setDatapath("C:\\Java_tesseract\\Tess4J\\tessdata");
        tesseract.setLanguage("rus");
        tesseract.setPageSegMode(1);
        tesseract.setOcrEngineMode(1);
        String strDoc = tesseract.doOCR(image);

        System.out.println(strDoc);

        String tempPlace = " ";
        String name = " ";
        String  place = " ";
        String surname = " ";
        String tempSurName = " ";
        String tempName = " ";
        String restemp = strDoc.replaceAll("\\s+", " ");
        String result = (restemp.trim()).toLowerCase(Locale.ROOT);

        String[] strArr = result.split(" ");

        for (int i1 = 0; i1 <strArr.length-1; i1++){
            if (strArr[i1].equals("имя")) {
                for (int i = i1; i < strArr.length - 1; i++) {
                    if (strArr[i + 1].equals("фамилия")) {
                        break; }
                   tempName = strArr[i + 1];
                    name = name+ " " + tempName; } } }

        for (int i1 = 0; i1 <strArr.length-1; i1++){
            if (strArr[i1].equals("фамилия")) {
                for (int i = i1; i < strArr.length - 1; i++) {
                    if (strArr[i + 1].equals("место")) {
                        break; }
                    tempSurName = strArr[i + 1];
                    surname = surname + " " + tempSurName; } } }

        for (int i1 = 0; i1 <strArr.length-1; i1++){
            if (strArr[i1].equals("жительства")) {
                for (int i = i1; i < strArr.length - 1; i++) {
                    if (strArr[i + 1].equals("")) {
                        break; }
                    tempPlace = strArr[i + 1];
                    place = place + " " + tempPlace; } } }


    HWPFDocument doc = null;
    try {FileInputStream fs = new FileInputStream("C:\\Users\\Waff\\Desktop\\picToDoc\\src\\docPass1.doc");
    doc = new HWPFDocument(fs);
    fs.close();
    }
    catch (Exception e){
    System.err.println("Error: document not found");
    }

    try {
        doc.getRange().replaceText("<name>", name);
        doc.getRange().replaceText("<surname>", surname);
        doc.getRange().replaceText("<place>", place);
    }
     catch (Exception e){
     System.err.println("Error replace text!");
     }
    FileOutputStream fo = new FileOutputStream("C:\\1.doc");
    doc.write(fo);
    fo.close();
    }
}