import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class PreencherPlanilha {

    public static final String FILA_PATH = "/home/roberto/Documentos/Relação_Produtos_Site.xlsx";
    public static final int NUMERO_LINHAS_COLUNA_A = 4457;
    public static final int NUMERO_LINHAS_COLUNA_J = 2138;
    
    public static void main(String[] args) throws IOException {

	try {

	    FileInputStream file = new FileInputStream(new File(FILA_PATH));

	    HSSFWorkbook workbook = new HSSFWorkbook(file);
	    HSSFSheet sheet = workbook.getSheetAt(0);

	    for (int i = 1; i < NUMERO_LINHAS_COLUNA_A; i++) {

		Row row = sheet.getRow(i);
		
		for (int i2 = 1; i2 < NUMERO_LINHAS_COLUNA_J; i2++) {
		
		    if (row.getCell(0).getNumericCellValue() == sheet.getRow(i2).getCell(9).getNumericCellValue()) {
			
			row.getCell(2).setCellValue(sheet.getRow(i2).getCell(9).getNumericCellValue());
			
			if (row.getCell(3) == null) {
			    Cell cell3 = row.createCell(3);
			    cell3.setCellValue(sheet.getRow(i2).getCell(10).getStringCellValue());
			} else {
			    row.getCell(3).setCellValue(sheet.getRow(i2).getCell(10).getStringCellValue());
			}
			
			Cell cell4 = row.getCell(4);
			if (cell4 == null) {
			    cell4 = row.createCell(4);
			}
			
			if (sheet.getRow(i2).getCell(11).getCellType() == Cell.CELL_TYPE_NUMERIC) {
			    cell4.setCellValue(sheet.getRow(i2).getCell(11).getNumericCellValue()); 
			} else if (sheet.getRow(i2).getCell(11).getCellType() == Cell.CELL_TYPE_STRING) {
			    cell4.setCellValue(sheet.getRow(i2).getCell(11).getStringCellValue());
			}
			
			Cell cell5 = row.getCell(5);
			if (cell5 == null) {
			    cell5 = row.createCell(5);
			}
			
			if (sheet.getRow(i2).getCell(12).getCellType() == Cell.CELL_TYPE_NUMERIC) {
			    cell5.setCellValue(sheet.getRow(i2).getCell(12).getNumericCellValue()); 
			} else if (sheet.getRow(i2).getCell(12).getCellType() == Cell.CELL_TYPE_STRING) {
			    cell5.setCellValue(sheet.getRow(i2).getCell(12).getStringCellValue());
			}
			
			Cell cell6 = row.getCell(6);
			if (cell6 == null) {
			    cell6 = row.createCell(6);
			}
			
			if (sheet.getRow(i2).getCell(13).getCellType() == Cell.CELL_TYPE_NUMERIC) {
			    cell6.setCellValue(sheet.getRow(i2).getCell(13).getNumericCellValue()); 
			} else if (sheet.getRow(i2).getCell(13).getCellType() == Cell.CELL_TYPE_STRING) {
			    cell6.setCellValue(sheet.getRow(i2).getCell(13).getStringCellValue());
			}
			
			Cell cell7 = row.getCell(7);
			if (cell7 == null) {
			    cell7 = row.createCell(7);
			}
			
			if (sheet.getRow(i2).getCell(14).getCellType() == Cell.CELL_TYPE_NUMERIC) {
			    cell7.setCellValue(sheet.getRow(i2).getCell(14).getNumericCellValue()); 
			} else if (sheet.getRow(i2).getCell(14).getCellType() == Cell.CELL_TYPE_STRING) {
			    cell7.setCellValue(sheet.getRow(i2).getCell(14).getStringCellValue());
			}
			
			break;
		    }
		}
	    }

	    file.close();

	    FileOutputStream outFile = new FileOutputStream(new File(FILA_PATH));
	    workbook.write(outFile);
	    outFile.close();
	    workbook.close();
	    System.out.println("Arquivo Excel editado com sucesso!");

	} catch (FileNotFoundException e) {
	    e.printStackTrace();
	    System.out.println("Arquivo Excel não encontrado!");
	} catch (IOException e) {
	    e.printStackTrace();
	    System.out.println("Erro na edição do arquivo!");
	}
    }
}
