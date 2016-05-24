
import java.util.*;
import java.util.regex.*;
import java.io.*;
import java.text.DateFormat;
import java.text.SimpleDateFormat;

import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;

public class Main {
	static LinkedHashMap<Integer,Pracownik> danePracownikow = new LinkedHashMap<>();
	static HashMap<Integer, LinkedList<Rozmowa>> polaczeniaPerFMNO = new HashMap<>();
	
	public static XSSFRow getMeRow (Iterator<Row> rowIterator){
		if (rowIterator.hasNext()) {		
			return (XSSFRow) rowIterator.next();
		} else
			return null;
	}
	
	public static void loadInvoice (String fname){
		int nrWierszaIn = 0;
		try (	FileInputStream strumienOdczytu = new FileInputStream(new File(fname));
				XSSFWorkbook workbook = new XSSFWorkbook(strumienOdczytu)	) {		
			XSSFSheet spreadsheet = workbook.getSheetAt(0);
			Iterator<Row> rowIterator = spreadsheet.iterator();	
			XSSFRow wiersz = null;

			while ((wiersz = getMeRow(rowIterator)) != null){
				nrWierszaIn++;
				if (wiersz.getCell(0).getCellType() == Cell.CELL_TYPE_NUMERIC)
					dodajRozmowe(odczytajRozmowe(wiersz));	
			}
		} catch (Exception ex){
			System.out.println("Problem przy przetwarzaniu faktury, numer wiersza: " + nrWierszaIn);
			ex.printStackTrace();
			System.exit(3);
		} 
		System.out.println("Zakończyłem przetwarzanie faktury");
	}
	
	public static Rozmowa odczytajRozmowe(XSSFRow wiersz) throws Exception{
		Cell komorka = null;
		Rozmowa call = null;
		DataFormatter formatter = new DataFormatter();
		SimpleDateFormat formatDaty = new SimpleDateFormat ("MM/dd/yy HH:mm:ss");
		Date rozmowaTimestamp = null;
		
		// przygotowanie komórek koszt i czas rozmowy do wczytania
		komorka = wiersz.getCell(10);
		String koszt = komorka.getStringCellValue().replace(",", ".");
		rozmowaTimestamp = formatDaty.parse(formatter.formatCellValue(wiersz.getCell(3)) + " " + formatter.formatCellValue(wiersz.getCell(4)));
		// public Rozmowa(String nrTelefonu, Date data, String kraj, String typ, String opis, String wybranyNumer, String czas, double koszt)
		call = new Rozmowa (
				formatter.formatCellValue(wiersz.getCell(2)),
				rozmowaTimestamp,
				wiersz.getCell(5).getStringCellValue(),
				wiersz.getCell(7).getStringCellValue(),
				formatter.formatCellValue(wiersz.getCell(6)),
				formatter.formatCellValue(wiersz.getCell(8)),
				formatter.formatCellValue(wiersz.getCell(9)),
				Double.parseDouble(koszt));
		
		// po załadowaniu danych z wiersza, trzeba jeszcze sprawdzić opis
		if (call.opis.length() < 2) call.opis = call.typ;
		return call;
	}
	
	public static void dodajRozmowe(Rozmowa call) {
		int numer = Integer.parseInt(call.getNrTelefonu());
		if (polaczeniaPerFMNO.containsKey(numer)) polaczeniaPerFMNO.get(numer).add(call);
		else {
			LinkedList<Rozmowa> nowaLista = new LinkedList<>();
			nowaLista.add(call);
			polaczeniaPerFMNO.put(numer, nowaLista);
		}
	}
	
	public static void loadVacationData (String fname){
		XSSFRow wiersz;
		Cell komorka;
		String ciag = null;
		Iterator<Row> rowIterator = null;
		// wczytywanie pliku fname
		try (	FileInputStream strumienOdczytu = new FileInputStream(new File(fname));
				XSSFWorkbook workbook = new XSSFWorkbook(strumienOdczytu)	) 		{	
			XSSFSheet spreadsheet = workbook.getSheetAt(0);
			rowIterator = spreadsheet.iterator();			
	
			// pętla zczytująca cały plik
			while ((wiersz = getMeRow(rowIterator)) != null){
				komorka = wiersz.getCell(0);
				
				// znalazłem String w pierwszej komórce danego wiersza
				if (komorka.getCellType() == Cell.CELL_TYPE_STRING) {
					ciag = komorka.getStringCellValue();
					Pattern p = Pattern.compile(".+Nr teczki (\\d+)");
					Matcher m = p.matcher(ciag);
					// czy wczytany String to użytkownik?
					if (m.find()){
						// znaleźliśmy użytkownika! Odczytujemy jego dane
						int numerFirmowy = Integer.parseInt(m.group(1));
						Pracownik user = danePracownikow.get(numerFirmowy);
						if(user != null) {
							// petla zczytująca urlopy, kończy się aż przeczyta komórkę "Razem"			
							boolean koniec = false;
							while (!koniec){
								wiersz = getMeRow(rowIterator);
								// czytam komórkę. jeśli to "Razem" kończę zczytywanie
								// jeśli to coś innego, znalazłem jakieś wolne
								if ((komorka = wiersz.getCell(2)) != null)
									try {
										if ((komorka.getCellType() == Cell.CELL_TYPE_STRING) && (komorka.getStringCellValue() != "")) {
											ciag = komorka.getStringCellValue();
											if (ciag.matches(".*Razem.*")) {
												koniec = true;
											} else {
												SimpleDateFormat ft = new SimpleDateFormat ("yyyy/MM/dd HH:mm:ss");
												Date dataStart, dataEnd;
												// kolejne dwie komórki mają datę rozpoczęcia i zakończenia wolnego
												komorka = wiersz.getCell(3);
												dataStart = ft.parse(komorka.getStringCellValue() + " 00:00:00");
												komorka = wiersz.getCell(4);
												dataEnd = ft.parse(komorka.getStringCellValue() + " 23:59:59");
												System.out.println("Pracownik: " + user.getName() + "(" + numerFirmowy + "), " + dataStart + " - " +dataEnd );
												Wolne urlop = new Wolne(dataStart, dataEnd);
												user.urlopy.add(urlop);
											}
										} else {
											System.out.println("Nie mogę rozpoznać wiersza dla pracownika " + numerFirmowy);
											System.exit(1);
										} 
									} catch (Exception ex) {
										System.out.println("Nr pracownika: " + numerFirmowy);
										ex.printStackTrace();
									}				
							}
						}
					}		
				}			
			}
		} catch (Exception ex){
			System.out.println("Wczytywanie plik nie powiodło się: " + fname);
			ex.printStackTrace();
			System.exit(1);
		}
	}
	
	public static void loadUserList(String fname){
		XSSFRow wiersz;
		DataFormatter formatter = new DataFormatter();
		Iterator<Row> rowIterator = null;
		
		try (	FileInputStream strumienOdczytu = new FileInputStream(new File(fname));
				XSSFWorkbook workbook = new XSSFWorkbook(strumienOdczytu)	) {		
			XSSFSheet spreadsheet = workbook.getSheetAt(0);
			rowIterator = spreadsheet.iterator();			
			
			while ((wiersz = getMeRow(rowIterator)) != null){
				String name, numery;
				int fmno = 0;
				// sprawdzamy czy pierwsza komórka wczytanego wiersza zawiera cyfry (FMNO)
				if ((wiersz.getCell(0) != null) && (wiersz.getCell(0).getCellType() == Cell.CELL_TYPE_NUMERIC)) {
					fmno = (int)wiersz.getCell(0).getNumericCellValue();
					name = wiersz.getCell(1).getStringCellValue();
					numery = formatter.formatCellValue(wiersz.getCell(4));
					Pracownik user = new Pracownik(name, fmno);
					danePracownikow.put(fmno, user);
					Pattern p = Pattern.compile("(\\d{9})");
					Matcher m = p.matcher(numery);
					while (m.find())
							user.numery.add(Integer.parseInt(m.group()));
					}
			}	
		} catch (Exception ex) {
			System.out.println("Wczytywanie listy użytkowników");
			ex.printStackTrace();
			System.exit(1);
		}
	}
	
	public static XSSFCellStyle createStyle(XSSFWorkbook wb, boolean bold, boolean red){
	    XSSFFont font = wb.createFont();
	    font.setFontHeightInPoints((short) 8);
	    font.setFontName("Arial");
	    if (red) font.setColor(IndexedColors.RED.getIndex());
	    font.setBold(bold);
	    XSSFCellStyle style = wb.createCellStyle();
	    style.setFont(font);
	    return style;
	}
	
	public static void generateFiles (){
		// Teraz pora skombinować ze sobą te wszystkie informacje.
		// dla każdego numeru użytkownika:
		//   dla każdego numeru telefonu tego użytkownika
		// 	   otworzyć plik
		//     wyciągnąć listę jego opłat
		//     wpisać tą listę do pliku uważając na urlopy
		//     zapisać podsumowanie
		
		verifyOutputFolder("wyniki");
		List<String> naglowek = generateHeader();
		
		for (Map.Entry<Integer, Pracownik> wpis : danePracownikow.entrySet()) {
			Pracownik pracownik = wpis.getValue();
			LinkedList<Integer> numery = (LinkedList<Integer>)pracownik.getNumery();
			
			for (int numer : numery) {
				double sumaDaneZaUrlopy = 0;
				double sumaRozmowyZaUrlopy = 0;
				double suma = 0;
				int iloscDanychTotal = 0;
				int iloscDanychRoaming = 0;
				
				LinkedList<Rozmowa> lista = polaczeniaPerFMNO.get(numer);
				if (lista != null) {
					//najpierw posortuję tą listę
					Collections.sort(lista);
					//otworzyć nowy plik dla danej kombinacji pracownik/numer
					XSSFWorkbook workbook = new XSSFWorkbook();
					DataFormat format = workbook.createDataFormat();
					// definicje stylów:
					// wszystkie style mają ten sam font definiowany w createStyle
					// 		nagłówek (center, bold, bottom border)
					XSSFCellStyle stylNaglowek = createStyle(workbook, true, false);
					stylNaglowek.setBorderBottom(CellStyle.BORDER_THICK);
					stylNaglowek.setAlignment(CellStyle.ALIGN_CENTER);
					// 		kolumna centrowana
					XSSFCellStyle stylDaneCenter = createStyle(workbook, false, false);
					stylDaneCenter.setAlignment(CellStyle.ALIGN_CENTER);
					//		kolumna do lewej
					XSSFCellStyle stylDaneLeft = createStyle(workbook, false, false);
					stylDaneLeft.setAlignment(CellStyle.ALIGN_LEFT);
					stylDaneLeft.setDataFormat(format.getFormat("@"));
					//		kolumna do prawej
					XSSFCellStyle stylDaneRight = createStyle(workbook, false, false);
					stylDaneRight.setAlignment(CellStyle.ALIGN_RIGHT);
					stylDaneRight.setDataFormat(format.getFormat("@"));
					//		data, center
					XSSFCellStyle stylDataCenter = createStyle(workbook, false, false);
					CreationHelper createHelper = workbook.getCreationHelper();
					stylDataCenter.setDataFormat(createHelper.createDataFormat().getFormat("mm/dd/yyyy"));
					stylDataCenter.setAlignment(CellStyle.ALIGN_CENTER);
					//		godzina, center
					XSSFCellStyle stylTimeCenter = createStyle(workbook, false, false);
					createHelper = workbook.getCreationHelper();
					stylTimeCenter.setDataFormat(createHelper.createDataFormat().getFormat("HH:mm"));
					stylTimeCenter.setAlignment(CellStyle.ALIGN_CENTER);
					//		numer w formacie #,###.##
					XSSFCellStyle stylLiczbaCenter = createStyle(workbook, false, false);
					stylLiczbaCenter.setAlignment(CellStyle.ALIGN_CENTER);
					stylLiczbaCenter.setDataFormat((short)4);
					//		numer w formacie ##, do prawej
					XSSFCellStyle stylLiczbaRight = createStyle(workbook, false, false);
					//stylLiczbaCenter.setAlignment(CellStyle.ALIGN_RIGHT);
					stylLiczbaRight.setDataFormat((short)1);
					
					// 		kolumna centrowana
					XSSFCellStyle stylREDDaneCenter = createStyle(workbook, false, true);
					stylREDDaneCenter.setAlignment(CellStyle.ALIGN_CENTER);
					//		kolumna do lewej
					XSSFCellStyle stylREDDaneLeft = createStyle(workbook, false, true);
					stylREDDaneLeft.setDataFormat(format.getFormat("@"));
					stylREDDaneLeft.setAlignment(CellStyle.ALIGN_LEFT);
					//		kolumna do prawej
					XSSFCellStyle stylREDDaneRight = createStyle(workbook, false, true);
					stylREDDaneRight.setDataFormat(format.getFormat("@"));
					stylREDDaneRight.setAlignment(CellStyle.ALIGN_RIGHT);
					//		data, center
					XSSFCellStyle stylREDDataCenter = createStyle(workbook, false, true);
					stylREDDataCenter.setDataFormat(createHelper.createDataFormat().getFormat("mm/dd/yyyy"));
					stylREDDataCenter.setAlignment(CellStyle.ALIGN_CENTER);
					//		numer w formacie #,###.##
					XSSFCellStyle stylREDLiczbaCenter = createStyle(workbook, false, true);
					stylREDLiczbaCenter.setAlignment(CellStyle.ALIGN_CENTER);
					stylREDLiczbaCenter.setDataFormat((short)4);
					//		godzina, center
					XSSFCellStyle stylREDTimeCenter = createStyle(workbook, false, true);
					stylREDTimeCenter.setDataFormat(createHelper.createDataFormat().getFormat("HH:mm"));
					stylREDTimeCenter.setAlignment(CellStyle.ALIGN_CENTER);
					//		numer w formacie ##, do prawej
					XSSFCellStyle stylREDLiczbaRight = createStyle(workbook, false, true);
					//stylLiczbaCenter.setAlignment(CellStyle.ALIGN_RIGHT);
					stylLiczbaRight.setDataFormat((short)1);
					
					XSSFSheet spreadsheet = workbook.createSheet(pracownik.getName());
					XSSFRow row;			
					// nagłówek
					row = spreadsheet.createRow(0);
					int komorkaId = 0;
					
					for (String tresc : naglowek){
						Cell komorka = row.createCell(komorkaId++);
						komorka.setCellValue(tresc);
						komorka.setCellStyle(stylNaglowek);					
					}
					
					int rowId = 2;

					// tworzymy nowe wiersz i zapisujemy z niego kolejne dane dla danego pracownika
					System.out.println(pracownik);
					for (Rozmowa rozmowa : lista)	{
						

						boolean czyWolne = czyPolaczeniePodczasUrlopu(rozmowa, pracownik);
						boolean czyCzerwone = !czyTransmisjaDanych(rozmowa) && czyWolne;
						row = spreadsheet.createRow(rowId);

						Cell komorka;

						// dodajemy kolejne kolumny:
						// numer telefonu
						komorka = row.createCell(0);
						if(czyCzerwone) komorka.setCellStyle(stylREDDaneCenter);
						else
							komorka.setCellStyle(stylDaneCenter);
						komorka.setCellValue(numer);

						
						// data
						komorka = row.createCell(1);
						DateFormat df = new SimpleDateFormat("MM/dd/yyyy");
						String data = df.format(rozmowa.data);
						komorka.setCellValue(data);
						if(czyCzerwone) komorka.setCellStyle(stylREDDataCenter);
						else
							komorka.setCellStyle(stylDataCenter);
						
						// godzina
						komorka = row.createCell(2);
						komorka.setCellValue(rozmowa.data);
						df = new SimpleDateFormat("HH:mm:ss");
						String godzina = df.format(rozmowa.data);
						komorka.setCellValue(godzina);
						if(czyCzerwone) komorka.setCellStyle(stylREDTimeCenter);
						else
							komorka.setCellStyle(stylTimeCenter);
						
						// kraj
						komorka = row.createCell(3);
						komorka.setCellValue(rozmowa.kraj);
						if(czyCzerwone) komorka.setCellStyle(stylREDDaneCenter);
						else
							komorka.setCellStyle(stylDaneCenter);
						
						// rodzaj połączenia
						komorka = row.createCell(4);
						komorka.setCellValue(rozmowa.typ);
						if(czyCzerwone) komorka.setCellStyle(stylREDDaneLeft);
						else
							komorka.setCellStyle(stylDaneLeft);

						// opis połączenia
						komorka = row.createCell(5);
						if(rozmowa.opis != "") komorka.setCellValue(rozmowa.opis);
						else komorka.setCellValue(rozmowa.opis);
						if(czyCzerwone) komorka.setCellStyle(stylREDDaneLeft);
						else
							komorka.setCellStyle(stylDaneLeft);
						
						// wybrany numer
						komorka = row.createCell(6);
						if (rozmowa.wybranyNumer.matches("[\\d]+")) {
							if(czyCzerwone) komorka.setCellStyle(stylREDDaneLeft);
							else komorka.setCellStyle(stylDaneLeft);
						} else {
							if(czyCzerwone) komorka.setCellStyle(stylREDDaneRight);
							else komorka.setCellStyle(stylDaneRight);	
						}
						komorka.setCellValue(rozmowa.wybranyNumer);
						
						// czas połączenia albo ilość pobranych danych
						komorka = row.createCell(7);
						if (rozmowa.czas.matches("[\\d]+")) {							
							if(czyCzerwone) komorka.setCellStyle(stylREDLiczbaRight);
							else komorka.setCellStyle(stylLiczbaRight);
							komorka.setCellValue(Long.parseLong(rozmowa.czas));
						} else {
							if(czyCzerwone) komorka.setCellStyle(stylREDDaneCenter);
							else komorka.setCellStyle(stylDaneCenter);	
							komorka.setCellValue(rozmowa.czas);
						}
						Pattern p = Pattern.compile("(\\d+) kB");
						Matcher m = p.matcher(rozmowa.czas);
						if(m.find()) {
							int dane = Integer.parseInt(m.group(1));
							iloscDanychTotal += dane;
							if(!rozmowa.kraj.matches("Polska")) iloscDanychRoaming += dane;	
						}
						
						
						// koszt połączenia
						komorka = row.createCell(8);
						komorka.setCellValue(rozmowa.koszt);
						if(czyCzerwone) {
							komorka.setCellStyle(stylREDLiczbaCenter);
						}
						else
							komorka.setCellStyle(stylLiczbaCenter);
						if(czyWolne) {
							if (rozmowa.typ.matches("Przesyłanie.*"))	sumaDaneZaUrlopy += rozmowa.koszt;
							else sumaRozmowyZaUrlopy += rozmowa.koszt;
						}
						suma += rozmowa.koszt;
						
						rowId++;	
						row.setHeight((short)230);
					}
					
					// i robimy SUMĘ kolumny koszt
					row = spreadsheet.createRow(1);
					Cell komorka = row.createCell(7);
					komorka.setCellStyle(stylNaglowek);
					komorka.setCellValue("SUMA:");
					komorka = row.createCell(8);
					komorka.setCellStyle(stylNaglowek);
					komorka.setCellFormula("SUM(I3:I"+rowId+")");
					row.setHeight((short)230);
					
					// wyliczami wartosci potrzebne do zestawienia i zapisujemy je w danych pracownika
					pracownik.setTotalCost(suma + pracownik.getTotalCost());
					pracownik.setPersonalDataCost(sumaDaneZaUrlopy + pracownik.getPersonalDataCost());
					pracownik.setPersonalCallCost(sumaRozmowyZaUrlopy + pracownik.getPersonalCallCost());
					pracownik.setDataTransferred(iloscDanychTotal + pracownik.getDataTransferred());
					pracownik.setDataRoaming(iloscDanychRoaming + pracownik.getDataRoaming());
					
					// autosize wszystkich kolumn
					for (int i = 0; i< 9; i++) {
						spreadsheet.autoSizeColumn(i);
						spreadsheet.setColumnWidth(i,spreadsheet.getColumnWidth(i) + 512);
					}
					// a teraz zapisujemy to do pliku
					try (FileOutputStream out = new FileOutputStream(new File("wyniki\\" + pracownik.getName() +"_"+ numer + ".xlsx"))){
						workbook.write(out);
					} catch (Exception ex) {
						System.out.println("Problem przy zapisie arkusza do pliku: " + pracownik.getName());
					}
				}
			}
		}
	}

	private static boolean czyPolaczeniePodczasUrlopu(Rozmowa call, Pracownik pracownik){
		boolean czyWolne = false;
		// sprawdzamy czy to urlop
		if (pracownik.urlopy != null)
			for (Wolne urlop : pracownik.urlopy) {
				if (!((call.data.compareTo(urlop.start) < 0)||(call.data.compareTo(urlop.end) > 0)))
					czyWolne = true;
		}
		return czyWolne;
	}
	
	private static boolean czyTransmisjaDanych(Rozmowa call){
		if (call.czas.matches("\\d+\\skB"))
			return true;
		else return false;
	}
	
	private static void verifyOutputFolder(String nazwaKatalogu){
		File katalog = new File (nazwaKatalogu);
		if(!katalog.exists()) {
			katalog.mkdirs();
		} else if (!katalog.isDirectory()) {
			katalog.delete();
			katalog.mkdirs();
		}
	}
	
	private static List<String> generateHeader(){
		LinkedList<String> naglowek = new LinkedList<>();		
		naglowek.add("Numer telefonu");
		naglowek.add("Data");
		naglowek.add("Godzina");
		naglowek.add("Kraj");
		naglowek.add("Typ połączenia");
		naglowek.add("Opis połączenia");
		naglowek.add("Wybrany numer");
		naglowek.add("Czas/kB/N");
		naglowek.add("Koszt");
		return naglowek;
	}
	
	public static void generateSummary(){
		
		// przygotuj workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet spreadsheet = workbook.createSheet();
		
		Cell komorka;
		XSSFRow row;
		
		// definicje stylów:
		// wszystkie style mają ten sam font definiowany w createStyle
		// 		nagłówek (center, bold, bottom border)
		XSSFCellStyle stylNaglowek = createStyle(workbook, true, false);
		stylNaglowek.setBorderBottom(CellStyle.BORDER_THICK);
		stylNaglowek.setAlignment(CellStyle.ALIGN_CENTER);
		// 		kolumna centrowana
		XSSFCellStyle stylDaneRight = createStyle(workbook, false, false);
		stylDaneRight.setAlignment(CellStyle.ALIGN_RIGHT);
		//		kolumna do lewej
		XSSFCellStyle stylDaneLeft = createStyle(workbook, false, false);
		stylDaneLeft.setAlignment(CellStyle.ALIGN_LEFT);
		//		numer w formacie ##, do prawej
		XSSFCellStyle stylLiczbaRight = createStyle(workbook, false, false);
		stylLiczbaRight.setDataFormat((short)164);
		stylLiczbaRight.setAlignment(CellStyle.ALIGN_RIGHT);
		//		numer w formacie #,##0.00\ [$zł-415], do prawej
		XSSFCellStyle stylREDLiczbaRight = createStyle(workbook, true, true);
		stylREDLiczbaRight.setDataFormat((short)164);
		stylREDLiczbaRight.setAlignment(CellStyle.ALIGN_RIGHT);
		//		numer w formacie #,###.##, Right
		XSSFCellStyle stylLiczbakBRight = createStyle(workbook, false, false);
		stylLiczbakBRight.setDataFormat((short)4);
		stylLiczbakBRight.setAlignment(CellStyle.ALIGN_RIGHT);
		
		row = spreadsheet.createRow(0);
		
		komorka = row.createCell(0);
		komorka.setCellStyle(stylNaglowek);
		komorka.setCellValue("Name");

		komorka = row.createCell(1);
		komorka.setCellStyle(stylNaglowek);
		komorka.setCellValue("Firm number");
		
		komorka = row.createCell(2);
		komorka.setCellStyle(stylNaglowek);
		komorka.setCellValue("Total mobile bill");

		komorka = row.createCell(3);
		komorka.setCellStyle(stylNaglowek);
		komorka.setCellValue("Personal call");
		
		komorka = row.createCell(4);
		komorka.setCellStyle(stylNaglowek);
		komorka.setCellValue("Personal data");
		
		komorka = row.createCell(5);
		komorka.setCellStyle(stylNaglowek);
		komorka.setCellValue("All data");	
		
		komorka = row.createCell(6);
		komorka.setCellStyle(stylNaglowek);
		komorka.setCellValue("Roaming data");
		
		int rowId = 1;
		for (Map.Entry<Integer, Pracownik> wpis : danePracownikow.entrySet()) {
			Pracownik pracownik = wpis.getValue();		
			row = spreadsheet.createRow(rowId);
			
			if(pracownik.getTotalCost() > 0) {
				// dodajemy kolejne kolumny:
				// Name, FMNO, Total Cost, Personal Cost
				komorka = row.createCell(0);
				komorka.setCellStyle(stylDaneLeft);
				komorka.setCellValue(pracownik.getName());
		
				komorka = row.createCell(1);
				komorka.setCellStyle(stylDaneRight);
				komorka.setCellValue(pracownik.getFMNO());
					
				komorka = row.createCell(2);
				komorka.setCellStyle(stylLiczbaRight);
				komorka.setCellValue(pracownik.getTotalCost());
		
				komorka = row.createCell(3);
				komorka.setCellStyle(stylREDLiczbaRight);
				komorka.setCellValue(pracownik.getPersonalCallCost());
					
				komorka = row.createCell(4);
				komorka.setCellStyle(stylREDLiczbaRight);
				komorka.setCellValue(pracownik.getPersonalDataCost());
				
				komorka = row.createCell(5);
				komorka.setCellStyle(stylLiczbakBRight);
				komorka.setCellValue(pracownik.getDataTransferred());
				
				komorka = row.createCell(6);
				komorka.setCellStyle(stylLiczbakBRight);
				komorka.setCellValue(pracownik.getDataRoaming());
				
				rowId++;	
			}
		}
		
		for (int i = 0; i< 7; i++) {
			spreadsheet.autoSizeColumn(i);
			spreadsheet.setColumnWidth(i,spreadsheet.getColumnWidth(i) + 512);
		}
		
		try (FileOutputStream out = new FileOutputStream(new File("wyniki\\Zestawienie.xlsx"))){
			workbook.write(out);
		} catch (Exception ex) {
			System.out.println("Problem przy zapisie arkusza do pliku: Zestawienie.xlsx");
		}
	}
	
	public static void main(String[] args) {
		
		System.out.println("Wczytuję plik z konfiguracją: orangeInvoice.ini");
		File plikIni = new File ("orangeInvoice.ini");
		String user_data = "Telefony.txt";
		String vacation_data = "Wolne.xlsx";
		String invoice = "Faktura.xlsx";
		
		if(plikIni.exists() && !plikIni.isDirectory()) {
			try {
				IniFile iniFile = new IniFile("orangeInvoice.ini");
				user_data = iniFile.getString("files", "phones", "Telefony.xlsx");
				vacation_data = iniFile.getString("files", "vacations", "Wolne.xlsx");
				invoice = iniFile.getString("files", "invoice", "Faktura.xlsx");
			} catch (Exception ex) {}
		}
		
		
		System.out.println("Przetwarzam listę użytkowników: " + user_data);
		loadUserList(user_data);
		
		System.out.println("Przetwarzam listę urlopową: " + vacation_data);
		loadVacationData(vacation_data);
		
		
		System.out.println("Przetwarzam fakturę: " + invoice);
		loadInvoice(invoice);
		
		System.out.println("Generuję pliki wynikowe.");
		generateFiles();
		
		System.out.println("Generuję zestawienie urlopowe.");
		generateSummary();
		
		System.out.println("Program zakończył działanie.");
	}

}
