
import java.util.*;
import java.io.*;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

public class TesterXLSX {

	public static XSSFRow getMeRow (Iterator<Row> rowIterator){
		if (rowIterator.hasNext()) {
			return (XSSFRow) rowIterator.next();
		} else
			return null;
	}
	
	public static void main(String[] args) {
		// TODO Auto-generated method stub
		XSSFWorkbook workbook = null;
		XSSFRow wiersz;
		Cell komorka;

		
		// wczytywanie pliku fname.xlsx
		File plik = new File("Test.xlsx");
		try (FileInputStream fIP = new FileInputStream(plik)) {
			
			workbook = new XSSFWorkbook(fIP);			
			XSSFSheet spreadsheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = spreadsheet.iterator();			

			
			// pętla zczytująca cały plik
			while ((wiersz = getMeRow(rowIterator)) != null){
				komorka = wiersz.getCell(0);
				System.out.println(komorka.toString());
				XSSFCellStyle styl = (XSSFCellStyle) komorka.getCellStyle();
				System.out.println(styl.getDataFormatString());
				System.out.println(styl.getDataFormat());
			}
			
		} catch (Exception ex){
			System.out.println("Wczytywanie plik: ");
			ex.printStackTrace();
			System.exit(1);
		}

	
	}

}
