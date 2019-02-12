import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;

import org.apache.log4j.BasicConfigurator;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import com.borland.silk.keyworddriven.annotations.Argument;

public class dataProcessing {

	protected static Logger logger = Logger.getLogger("");

	@Test
	// @Keyword(value = "Sauvegarder des donnees dans un fichier Excel")
	public void SaveData(@Argument("Numéro Opération à sauvegarder") String operationNum,
			@Argument("CheminFichier") String filePath, @Argument("Nom du feuille Excel") String sheetName, int ligne,
			int Cellule) throws IOException {
		BasicConfigurator.configure();
		FileReader fileReader = null;
		try {
			fileReader = new FileReader(filePath);
		} catch (FileNotFoundException e) {
			logger.error("erreur", e);
		} finally {
			if (fileReader != null) {
				try {
					fileReader.close();
				} catch (IOException e) {
					logger.error("Erreur lecture du fichier", e);
				}
			}
		}

		FileInputStream file = new FileInputStream(filePath);
//        
		Workbook workbook = new XSSFWorkbook(file);
		int sheetNum = workbook.getSheetIndex(sheetName);
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetNum);
		int lastRow = sheet.getLastRowNum();
		System.out.println("ligne = " + lastRow);

		XSSFRow row = (XSSFRow) sheet.getRow(ligne); // 1 : ligne 1
		Cell cellule = row.getCell(Cellule); // colonne 21

		try {
			cellule.setCellValue(operationNum);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		file.close();
		FileOutputStream outFile = new FileOutputStream(filePath);
		workbook.write(outFile);
		workbook.close();
		outFile.close();
		logger.info("Le stockage de numéro d'operation dans le fichier Excel des données est fait");
		BasicConfigurator.resetConfiguration();
	}

	// @Keyword(value = "Lire des donnees du fichier Excel")
	public String ReadData(@Argument("CheminFichier") String filePath,
			@Argument("Nom du feuille Excel") String sheetName, int ligne, int Cellule) throws IOException {
		BasicConfigurator.configure();
		FileReader fileReader = null;
		try {
			fileReader = new FileReader(filePath);
		} catch (FileNotFoundException e) {
			logger.error("erreur", e);
		} finally {
			if (fileReader != null) {
				try {
					fileReader.close();
				} catch (IOException e) {
					logger.error("Erreur lecture du fichier", e);
				}
			}
		}

		FileInputStream file = new FileInputStream(filePath);
//        
		Workbook workbook = new XSSFWorkbook(file);
		int sheetNum = workbook.getSheetIndex(sheetName);
		XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetNum);
		int lastRow = sheet.getLastRowNum();
		System.out.println("ligne = " + lastRow);

		XSSFRow row = (XSSFRow) sheet.getRow(ligne); // 1 : ligne 1
		Cell cellule = row.getCell(Cellule); // colonne 21
		try {
			String num = cellule.getStringCellValue();
			file.close();
			FileOutputStream outFile = new FileOutputStream(filePath);

			workbook.write(outFile);
			workbook.close();
			outFile.close();
			logger.info("La lecture de numéro d'operation du fichier Excel des données est fait :" + num);
			BasicConfigurator.resetConfiguration();
			return num;
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			String status = "ko";
			return status;
		}
		
	}

	public static void main(String[] args) throws IOException {
		dataProcessing dataProcessing = new dataProcessing();
		try {
			dataProcessing.SaveData("157584889946",
					"C:\\Users\\bha\\Desktop\\07122018pole-qualite-fonctionnel Compte rendu  de la rétrospective -sprint 9.xlsx",
					"Feuil1", 1, 21);
			String opereation = dataProcessing.ReadData(
					"C:\\Users\\bha\\Desktop\\07122018pole-qualite-fonctionnel Compte rendu  de la rétrospective -sprint 9.xlsx",
					"Feuil1", 1, 21);
			System.out.println(opereation);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			
		}
	}
}