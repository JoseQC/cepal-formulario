import io.reactivex.Observable;
import lombok.AllArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;
import static java.util.stream.Collectors.*;

@AllArgsConstructor
public class Excel {

    private static String[] columns = {"Familia", "Proc / Reu / Eve", "Plantilla", "Variable", "JooScript"};
    List<RowData> listVariable = new ArrayList<>();
    List<RowData> listVariableTotal = new ArrayList<>();
    List<RowData> listVariableTotalProcess = new ArrayList<>();
    List<RowData> listVariableExc = new ArrayList<>();
    Process process = new Process();
    List<String> listTypeReuEve =process.getListaVariables();
    List<String> listVariEvi = new ArrayList<>();
    List<String> listVariableHori=Arrays.asList("_AUXILIAR",
            "_PLANT",
            "_AYTO",
            "PIE_",
            "TAB_",
            "_label",
            "REGISTRO",
            "SISTEMA",
            "eEMC",
            "EELL",
            "_SOLIC");

    public Excel(List<String> listTypeReuEve) {
        this.listTypeReuEve = listTypeReuEve;
    }

    public  void readFileInfVar() throws IOException {

        File myFile = new File("C:\\Users\\Jose\\Workspace - Developer\\Java\\informe_plantillas.xlsx");
        FileInputStream fis = new FileInputStream(myFile);

        // Finds the workbook instance for XLSX file
        XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);

        // Return first sheet from the XLSX workbook
        XSSFSheet mySheet= myWorkBook.getSheet("RAW_DATA");
        //XSSFRow row = mySheet.getRow(0);

        Observable
                .fromIterable(mySheet)
                .filter(row -> {

                    RowData rowData = new RowData()
                            .family(row.getCell(0).getStringCellValue())
                            .typeProReuEve(row.getCell(1).getStringCellValue())
                            .template(row.getCell(2).getStringCellValue())
                            .variable(row.getCell(3).getStringCellValue())
                            .jooScript(Boolean.parseBoolean(row.getCell(4).getStringCellValue()));
                    listVariableTotal.add(rowData);
                    if(listTypeReuEve.contains(row.getCell(1).getStringCellValue()))
                    {

                        if(listVariEvi.contains(row.getCell(3).getStringCellValue())
                                || verified(row.getCell(3).getStringCellValue(),row.getCell(2).getStringCellValue()))
                        {
                           listVariableExc.add(rowData);
                        }
                        else {

                            listVariable.add(rowData);
                        }
                        listVariableTotalProcess.add(rowData);
                        return true;
                    }

                    else{

                        return false;
                    }


                })
                .subscribe();
        System.out.println("Total size: "+listVariableTotal.size());
        System.out.println("Total process size: "+listVariableTotalProcess.size());
        System.out.println("Filter size: "+listVariable.size());
        System.out.println("Filter exc size:"+listVariableExc.size() );

    }

    public boolean verified(String var, String plant)    {
        String data1="";
        for (String data : listVariableHori) {
            data1 = ".*" + data + ".*";
            if (!var.matches(data1)) {
                continue;
            }
            return true;
        }
      return false;
    }

    public void readVarEvi() throws IOException {

        File myFile = new File("C:\\Users\\Jose\\Workspace - Developer\\Java\\Variables a evitar.xlsx");
        FileInputStream fis = new FileInputStream(myFile);

        // Finds the workbook instance for XLSX file
        XSSFWorkbook myWorkBook = new XSSFWorkbook(fis);

        // Return first sheet from the XLSX workbook
        XSSFSheet mySheet= myWorkBook.getSheet("Hoja1");
        //XSSFRow row = mySheet.getRow(0);

        Observable
                .fromIterable(mySheet)
                .map(row -> {
                    return  listVariEvi.add(row.getCell(0).getStringCellValue());})
                .subscribe();
        System.out.println("size var evi"+listVariEvi.size());

    }

    public void createDataTotal() throws IOException {
        // Create a Workbook
        File myFile = new File("C:\\Users\\Jose\\Workspace - Developer\\Java\\informe_plantillas.xlsx");
        FileInputStream fis = new FileInputStream(myFile);

        // Finds the workbook instance for XLSX file
        XSSFWorkbook workbook = new XSSFWorkbook(fis);


        /* CreationHelper helps us create instances of various things like DataFormat,
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet("Data Procces Total");

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Create cells
        for(int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }
        // Create Other rows and cells with employees data
        int rowNum = 1;
        for(RowData data: listVariableTotalProcess) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0)
                    .setCellValue(data.family());

            row.createCell(1)
                    .setCellValue(data.typeProReuEve());

            row.createCell(2)
                    .setCellValue(data.template());

            row.createCell(3)
                    .setCellValue(data.variable());

            row.createCell(4)
                    .setCellValue(data.jooScript());
        }

        // Resize all columns to fit the content size
        for(int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(myFile);
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
    }

    public void createDataFilter() throws IOException {
        // Create a Workbook
        File myFile = new File("C:\\Users\\Jose\\Workspace - Developer\\Java\\informe_plantillas.xlsx");
        FileInputStream fis = new FileInputStream(myFile);

        // Finds the workbook instance for XLSX file
        XSSFWorkbook workbook = new XSSFWorkbook(fis);


        /* CreationHelper helps us create instances of various things like DataFormat,
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet("Data Filter Process");

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Create cells
        for(int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Create Cell Style for formatting Date
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

        // Create Other rows and cells with employees data
        int rowNum = 1;
        for(RowData data: listVariable) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0)
                    .setCellValue(data.family());

            row.createCell(1)
                    .setCellValue(data.typeProReuEve());

            row.createCell(2)
                    .setCellValue(data.template());

            row.createCell(3)
                    .setCellValue(data.variable());

            row.createCell(4)
                    .setCellValue(data.jooScript());
        }

        // Resize all columns to fit the content size
        for(int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(myFile);
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
    }

    public void createDataFilterExc() throws IOException {
        // Create a Workbook
        File myFile = new File("C:\\Users\\Jose\\Workspace - Developer\\Java\\informe_plantillas.xlsx");
        FileInputStream fis = new FileInputStream(myFile);

        // Finds the workbook instance for XLSX file
        XSSFWorkbook workbook = new XSSFWorkbook(fis);


        /* CreationHelper helps us create instances of various things like DataFormat,
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet("Data process Excluido");

        // Create a Font for styling header cells
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());

        // Create a CellStyle with the font
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);

        // Create a Row
        Row headerRow = sheet.createRow(0);

        // Create cells
        for(int i = 0; i < columns.length; i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(columns[i]);
            cell.setCellStyle(headerCellStyle);
        }

        // Create Cell Style for formatting Date
        CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

        // Create Other rows and cells with employees data
        int rowNum = 1;
        for(RowData data: listVariableExc) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0)
                    .setCellValue(data.family());

            row.createCell(1)
                    .setCellValue(data.typeProReuEve());

            row.createCell(2)
                    .setCellValue(data.template());

            row.createCell(3)
                    .setCellValue(data.variable());

            row.createCell(4)
                    .setCellValue(data.jooScript());
        }

        // Resize all columns to fit the content size
        for(int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream(myFile);
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();
    }

    public void getDuplicates() {

        List<RowData> listaUnica = getDuplicatesMap(listVariable).values().stream()
                .filter(duplicates -> duplicates.size() == 1)
                .flatMap(Collection::stream)
                .collect(toList());

        List<RowData> listaDupli = getDuplicatesMap(listVariable).values().stream()
                .filter(duplicates -> duplicates.size() > 1)
                .flatMap(Collection::stream)
                .collect(toList());

        Map<String, List<RowData>> duplicatesMap = getDuplicatesMap(listaDupli);
        Map<String, List<RowData>> duplicatesMapPrueb = getDuplicatesMap(listVariable);
        Map<String, List<RowData>> duplicatesMapPrueb1;

        for ( String text  :duplicatesMapPrueb.keySet()) {

            System.out.println("Variable: "+text );
            duplicatesMapPrueb.get(text).stream().forEach(r->System.out.println(r.toString()));

           /* if(duplicatesMapPrueb.get(text).size()>1){
                duplicatesMapPrueb1 = getDUplicateExtra(duplicatesMapPrueb.get(text));
                for(String text1 : duplicatesMapPrueb1.keySet()){
                        System.out.println("pruebaa:"+text1 + " otroo "+ duplicatesMapPrueb1.get(text1).size());
                        duplicatesMapPrueb1.get(text1).stream().forEach(r->System.out.println(r.toString()));
                }
            }*/
        }

        System.out.println("=====================================");
        //listaDupli.stream().forEach(r->System.out.println("Type: "+ r.typeProReuEve()+" Variable "+ r.variable()));

        System.out.println("Lista de duplicados tamaño:"+ listaDupli.size());
        System.out.println ("Lista de duplicados tamaño:"+ listaUnica.size());
    }

    private static Map<String, List<RowData>> getDuplicatesMap(List<RowData> personList) {
        return personList.stream().collect(groupingBy(RowData::variable ));

    }

    private static Map<String, List<RowData>> getDUplicateExtra(List<RowData> personList) {
        return personList.stream().collect(groupingBy(RowData::typeProReuEve));
    }
}
