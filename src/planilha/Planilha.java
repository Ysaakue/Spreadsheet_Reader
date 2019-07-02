/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package planilha;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author isaac
 */
public class Planilha {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws Exception {
       
        String FILE_NAME = "planilha.xlsx";  //caminho para os aquivo
    
        File arquivoExcel = new File(FILE_NAME); //abre o aquivo
        
        FileInputStream excelFile = new FileInputStream(arquivoExcel);
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet datatypeSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = datatypeSheet.iterator();
            
        iterator.next(); // pula a primeira linha, pois contém apenas títulos
        while (iterator.hasNext()) { //enquanto houver uma proxima linha pra ler executa
            Row currentRow = iterator.next();

            String curso      = currentRow.getCell(0).getStringCellValue();//guarda valor da primeira celula da linha
            String codigo_string = "";
            double codigo_double = -1000;
            
            try {
                if ( currentRow.getCell(1).getCellType() == CellType.STRING) { //verifica se o tipo da segunda celula é uma string
                    codigo_string = currentRow.getCell(1).getStringCellValue(); //guarda o valor da string da celula dois
                    System.out.println("Curso.create(nome:\""+ curso + "\", codigo: \"" + codigo_string + "\")");
                } else { //caso não seja string
                    codigo_double = currentRow.getCell(1).getNumericCellValue(); //guarda o valor do numeor na celula dois
                    System.out.println("Curso.create(nome:\""+ curso + "\", codigo: \"" + codigo_double + "\")");
                }
            } catch (Exception e) { //caso aconteça um erro
                System.out.println("Curso.create(nome:\""+ curso + "\", codigo: \"" + "\")"); //deixa o campo da celula dois me branco
            }
                     
        }
        
    } // fim do main
    
} // fim da classe
