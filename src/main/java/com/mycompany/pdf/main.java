/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.pdf;
import com.aspose.words.Document;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;


/**
 *
 * @author IDitroyer
 */
public class main {
     public static void main(String[] args) throws InterruptedException, Exception{
        String filePath = "c:\\Users\\IDitroyer\\Documents\\FORMATO E -300  Declaracion Privada Estampilla.doc";
        String filePath2 = "c:\\Users\\IDitroyer\\Documents\\FORMATO E -300  Declaracion Privada Estampilla2.doc";
        POIFSFileSystem fs;        
        try {            
            fs = new POIFSFileSystem(new FileInputStream(filePath));            
            HWPFDocument doc = new HWPFDocument(fs);
            doc = replaceText(doc, "«¬*VIGENCIA*¬»", "2016-2019");
            doc = replaceText(doc, "«¬*FECHAPRESENTACION*¬»", "25/07/2019");
            doc = replaceText(doc, "«¬*CONSECUTIVO*¬»", "20193010100000004");
            doc = replaceText(doc, "«¬*NIT*¬»", "BIMENSUAL");
            doc = replaceText(doc, "«¬*NOMBRE*¬»", "Jesús Ernesto");
            doc = replaceText(doc, "«¬*APELLIDO*¬»", "Quinto Narvaéz");
            doc = replaceText(doc, "«¬*RAZONSOCIAL*¬»", "MUCHAS RAZONES");
            doc = replaceText(doc, "«¬*DIRECCION*¬»", "CRA 18 N 63C 213");
            doc = replaceText(doc, "«¬*CIUDAD*¬»", "BARRANQUILLA");
            doc = replaceText(doc, "«¬*CORREO*¬»", "JESUSQUINTON@HOTMAIL.COM");
            doc = replaceText(doc, "«¬*TIPOID*¬»", "CC");
            doc = replaceText(doc, "«¬*ID*¬»", "1002182720");
            doc = replaceText(doc, "«¬*NOCONTRATO*¬»", "1002458d");
            doc = replaceText(doc, "«¬*CONTRATANTE*¬»", "Alumbrado publico");
            doc = replaceText(doc, "«¬*FECHACONTRATO*¬»", "17/07/2019");
            doc = replaceText(doc, "«¬*ENTIDAD*¬»", "Gobernación del Atlantico");
            doc = replaceText(doc, "«¬*MUNICIPIO*¬»", "Barranquilla");
            doc = replaceText(doc, "«¬*BV*¬»", "12.000.000.00");
            doc = replaceText(doc, "«¬*VE*¬»", "524.000.00");
            doc = replaceText(doc, "«¬*SC*¬»", "524.000.00");
            doc = replaceText(doc, "«¬*NP*¬»", "524.000.00");
            doc = replaceText(doc, "CODIGOBARRAS12345", "20193010100000004");
            
            
            
            
         

               saveWord(filePath2, doc);
            
               
               
               Document doc2 = new Document(filePath2);
               doc2.save("c:\\Users\\IDitroyer\\Documents\\FORMATO E -300  Declaracion Privada Estampilla3.pdf");

            
            
            
            
            
            
            
                
            
         
        }
        catch(FileNotFoundException e){
            e.printStackTrace();
        }
        catch(IOException e){
            e.printStackTrace();
        }
    }

    private static HWPFDocument replaceText(HWPFDocument doc, String findText, String replaceText){
        Range r1 = doc.getRange(); 

        for (int i = 0; i < r1.numSections(); ++i ) { 
            Section s = r1.getSection(i); 
            for (int x = 0; x < s.numParagraphs(); x++) { 
                Paragraph p = s.getParagraph(x); 
                for (int z = 0; z < p.numCharacterRuns(); z++) { 
                    CharacterRun run = p.getCharacterRun(z); 
                    String text = run.text();
                    if(text.contains(findText)) {
                        
                        run.replaceText(findText, replaceText);
                    } 
                }
            }
        } 
        return doc;
    }

    private static void saveWord(String filePath, HWPFDocument doc) throws FileNotFoundException, IOException{
        FileOutputStream out = null;
        try{
            out = new FileOutputStream(filePath);
            doc.write(out);
        }
        finally{
            out.close();
        }
    }
    
}
