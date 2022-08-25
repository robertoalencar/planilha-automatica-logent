import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;

public class IdentificarEmailsDuplicados {

    public static final String FILA_PATH = "/home/roberto/Documentos/Newsletter_3B_Betinho.xls";
    public static final int NUMERO_LINHAS_COLUNA_B = 19406;
    
    public static void main(String[] args) throws IOException {

	try {

	    FileInputStream file = new FileInputStream(new File(FILA_PATH));

	    HSSFWorkbook workbook = new HSSFWorkbook(file);
	    HSSFSheet sheet = workbook.getSheetAt(0);

	    for (int i = 1; i < NUMERO_LINHAS_COLUNA_B; i++) {

		Row row1 = sheet.getRow(i);
		Row row2 = sheet.getRow(i+1);

		if (row1 != null
			&& row1.getCell(1) != null 
			&& row1.getCell(1).getStringCellValue() != null
			&& row2 != null
			&& row2.getCell(1) != null
			&& row2.getCell(1).getStringCellValue() != null
			&& row1.getCell(1).getStringCellValue().trim().equals(row2.getCell(1).getStringCellValue().trim())) {
		    
		    row2.getCell(2).setCellValue(1);
		    
		} else {
		    
		    if (row2 != null
			    && row2.getCell(2) != null) {
			
			row2.getCell(2).setCellValue("");
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
