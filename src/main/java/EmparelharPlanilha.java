import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class EmparelharPlanilha {

    public static final String FILA_PATH = "/home/roberto/Documentos/Relacao_Produtos_Site2.xls";
    public static final int NUMERO_LINHAS_COLUNA_A = 3994;
    public static final int NUMERO_LINHAS_COLUNA_B = 2124;
    
    public static void main(String[] args) throws IOException {

	try {

	    FileInputStream file = new FileInputStream(new File(FILA_PATH));

	    HSSFWorkbook workbook = new HSSFWorkbook(file);
	    HSSFSheet sheet = workbook.getSheetAt(0);
	    
	    String[] listaCol1 = new String[NUMERO_LINHAS_COLUNA_A];
	    String[] listaCol2 = new String[NUMERO_LINHAS_COLUNA_A];
	    int ind = 0;

	    for (int i = 1; i < NUMERO_LINHAS_COLUNA_A; i++) {

		Row row = sheet.getRow(i);
		
		String valorColuna01 = obterValorColuna(row, 0);
		String valorColuna02 = obterValorColuna(row, 1);

		listaCol1[ind] = valorColuna01;
		listaCol2[ind] = valorColuna02;
		
		ind++;
	    }
	    
	    int linha = 1;
	    Row row;
	    
	    for (int i = 0; i < listaCol1.length; i++) {
		
		row = sheet.getRow(linha);
		if (row != null) {
		    row.getCell(3).setCellValue(listaCol1[i]);
		}
		
		for (int i2 = 1; i2 < listaCol2.length; i2++) {
		    
		    if ((listaCol1[i] != null && listaCol2[i2] != null) && (listaCol1[i].equals(listaCol2[i2]))) {
			
			if (row != null) {
			    row.getCell(4).setCellValue(listaCol2[i2]);
			}
		    }
		}
		
		linha++;
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
    
    public static String obterValorColuna(Row linha, int coluna) {

	String valorColuna = "";

	if (linha != null && linha.getCell(coluna) != null) {
	    
	    if (linha.getCell(coluna).getCellType() == Cell.CELL_TYPE_STRING) {
		
		valorColuna = linha.getCell(coluna).getStringCellValue();
		
	    } else if (linha.getCell(coluna).getCellType() == Cell.CELL_TYPE_NUMERIC) {
		
		int valor = (int) linha.getCell(coluna).getNumericCellValue();
		valorColuna = String.valueOf(valor);
	    }
	}
	
	if (valorColuna.length() == 1) {
	    valorColuna = "0000" + valorColuna;
	} else if (valorColuna.length() == 2) {
	    valorColuna = "000" + valorColuna;
	} else if (valorColuna.length() == 3) {
	    valorColuna = "00" + valorColuna;
	} else if (valorColuna.length() == 4) {
	    valorColuna = "0" + valorColuna;
	}

	return valorColuna;
    }

}
