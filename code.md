public class Exceleyazma {

   
    public static void main(String[] args) throws FileNotFoundException, IOException, EncryptedDocumentException, InvalidFormatException {
      
                
     InputStream inp = new FileInputStream("C:\\Users\\sonerc\\Desktop\\java dosyalarým kodlarým\\sonuc.xlsx");

    Workbook wb = WorkbookFactory.create(inp);
    Sheet sheet = wb.getSheetAt(0);

    
    Row row = sheet.createRow((short)sheet.getLastRowNum()+1);
    //sheet.shiftRows(1, sheet.getLastRowNum(), 1, true,true);
    
    row.createCell(0).setCellValue("A");
    row.createCell(1).setCellValue("B");
    row.createCell(2).setCellValue("This is a string");
    row.createCell(3).setCellValue(true);

    // Write the output to a file
    FileOutputStream fileOut = new FileOutputStream("C:\\Users\\sonerc\\Desktop\\java dosyalarým kodlarým\\sonuc.xlsx");
    wb.write(fileOut);
    fileOut.close();
    }
}
