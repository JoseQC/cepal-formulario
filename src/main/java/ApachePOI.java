
import com.poiji.bind.Poiji;
import com.poiji.option.PoijiOptions;
import io.reactivex.Observable;
import io.reactivex.functions.Function;
import io.reactivex.functions.Predicate;
import org.apache.poi.sl.extractor.SlideShowExtractor;
import org.apache.poi.sl.usermodel.SlideShow;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import reactor.core.publisher.Flux;
import reactor.core.publisher.Mono;

import java.io.*;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;
import java.util.stream.Collectors;

import static java.util.Arrays.asList;


public class ApachePOI {

    public static List<EventReutilizable> listEventReu = new ArrayList<>();
    public static List<String> listVariEvit = new ArrayList<>();
    public static boolean existEvents=false;
    public static boolean existReutilizable=false;
    public  static List<RowData> listRowData ;
    public  static List<RowData> listRowDataFilter = new ArrayList<>();
    public  static List<DataDescription> listDataDescription;
    public  static PoijiOptions options = PoijiOptions.PoijiOptionsBuilder.settings().build();
    public static String process = "PUB_LPP";
    public static String[] columns = {"Familia", "Proc / Reu / Eve", "Plantilla", "Variable","JooScript","Orden"};
    public static List<String> noti = asList("EVE_EVE_NOA_Notificacion_acuerdo",
            "EVE_EVE_NOR_Notificacion_resolucion",
            "MOD_REU_NOA_Notificacion_acuerdo",
            "MOD_REU_NOR_Notificacion_resolucion");
    public static List<String> filters=new ArrayList<>();

    public static void main(String...arg ) throws IOException {

        listEventReu.add(new EventReutilizable()
                .code(process)
                .name("Solicitud de permiso de tenencia de animales potencialmente peligrosos"));

        //PoijiOptions options = PoijiOptions.PoijiOptionsBuilder.settings().build();
        listRowData = Poiji.fromExcel(new File("x|cuments\\cep@l\\informe_plantillas.xlsx"), RowData.class, options);

        variablesEvitar();
        eventReutilizable();
        addOtherEventsAndReutilizable();
        listEventReu.forEach(o->System.out.println(""+o.toString()));
        getMetaDataPptx();
        filter();
        System.out.println("==============================");
        listRowDataFilter.forEach(System.out::println);
        System.out.println("==============================");
        getVisivilityDatosExpAndNoti();
        writeData();
        //order();
    }

    public static void getVisivilityDatosExpAndNoti(){


        Flux.fromIterable(listRowDataFilter)
                .distinct(RowData::getPlantilla).subscribe(r->System.out.println(r.getPlantilla()));


    }
    public static <T> Predicate<T> distinctByKey(Function<? super T, ?> keyExtractor) {
        Set<Object> seen = ConcurrentHashMap.newKeySet();
        return t -> seen.add(keyExtractor.apply(t));
    }

    public static void getMetaDataPptx() throws IOException {
        SlideShow<XSLFShape,XSLFTextParagraph> slideshow
                = new XMLSlideShow(new FileInputStream("C:\\Users\\jquipsec\\Documents\\cep@l\\7.10.Servicios p??blicos\\PUB_LPP\\PUB_LPP.pptx"));

        SlideShowExtractor<XSLFShape,XSLFTextParagraph> slideShowExtractor
                = new SlideShowExtractor<XSLFShape,XSLFTextParagraph>(slideshow);
        slideShowExtractor.setCommentsByDefault(true);
        slideShowExtractor.setMasterByDefault(true);
        slideShowExtractor.setNotesByDefault(true);

        String allTextContentInSlideShow = slideShowExtractor.getText();

        boolean solicitud = allTextContentInSlideShow.contains("MOD_REU_NOA_Notificacion_acuerdo");



        for(String str : noti)
        {
            if (!allTextContentInSlideShow.contains(str))
            {
                filters.add(str);
            }
        }

        System.out.println("size:"+filters);




    }

    private static void writeData() throws IOException {
        // Create a Workbook
        Workbook workbook = new XSSFWorkbook(); // new HSSFWorkbook() for generating `.xls` file

        /* CreationHelper helps us create instances of various things like DataFormat,
           Hyperlink, RichTextString etc, in a format (HSSF, XSSF) independent way */
        CreationHelper createHelper = workbook.getCreationHelper();

        // Create a Sheet
        Sheet sheet = workbook.createSheet("Filter");

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
        for(RowData rowData: listRowDataFilter) {
            Row row = sheet.createRow(rowNum++);

            row.createCell(0)
                    .setCellValue(rowData.getFamily());

            row.createCell(1)
                    .setCellValue(rowData.getProcReuEve());
            row.createCell(2)
                    .setCellValue(rowData.getPlantilla());
            row.createCell(3)
                    .setCellValue(rowData.getVariable());
            row.createCell(4)
                    .setCellValue(rowData.getJooScript());

        }

        // Resize all columns to fit the content size
        for(int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);
        }

        // Write the output to a file
        FileOutputStream fileOut = new FileOutputStream("filter.xlsx");
        workbook.write(fileOut);
        fileOut.close();

        // Closing the workbook
        workbook.close();

    }

    private static void variablesEvitar() throws IOException {


        File file = new File("C:\\Users\\jquipsec\\Documents\\Workspace - Jose\\Developer\\Java\\Cep@l\\cepal-formulario\\Variables a evitar.xlsx");
        FileInputStream inputStream = new FileInputStream(new File(String.valueOf(file)));

        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = firstSheet.iterator();

        while (iterator.hasNext()) {
            Row nextRow = iterator.next();
            Iterator<Cell> cellIterator = nextRow.cellIterator();
            Cell cell = cellIterator.next();
            listVariEvit.add(cell.getStringCellValue());
        }

        workbook.close();
        inputStream.close();
    }

    private static  void readTemplates() {

        //set path of file
        File file = new File("C:\\Users\\jquipsec\\Documents\\cep@l\\7.10.Servicios p??blicos\\PUB_LPP\\PUB_LPP");
        String[] list = file.list();

        listRowData = Poiji.fromExcel(new File("C:\\Users\\jquipsec\\Desktop\\informe_plantillas.xlsx"), RowData.class, options);

        //RxJava
        Observable.fromIterable(listRowData).map(r -> {
            assert list != null;
            for (String data : list) {
                String[] arrOfStr = data.split("\\.");
                if (r.getPlantilla().endsWith(arrOfStr[1])) {
                    r.setOrder(Integer.parseInt(arrOfStr[0]));
                }
            }
            return r;
        })
                .subscribe(t->System.out.println(t.getFamily()+", "+t.getProcReuEve()+", "+t.getPlantilla()+", "+t.getVariable()+", "+t.getJooScript()+", "+t.getOrder()));

    }

    private static void extractDescription(){

        listDataDescription = Poiji.fromExcel(new File("C:\\Users\\jquipsec\\Documents\\cep@l\\Diccionario_variables.xlsx"), DataDescription.class, options);
        //System.out.print(listDataDescription.size());
    }

    private static void  eventReutilizable() throws IOException {

        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("C:\\Users\\jquipsec\\Documents\\cep@l\\7.10.Servicios p??blicos\\PUB_LPP\\PUB_LPP.pptx"));
        Observable.fromIterable(ppt.getSlides())
                .map(a->Observable
                        .fromIterable(a.getShapes())
                        .filter(b->b instanceof XSLFGroupShape)
                        .map(y->(XSLFGroupShape)y)
                        .map(b->Observable
                                .fromIterable(b.getShapes())
                                .filter(f->f instanceof XSLFTextShape)
                                .map( t -> (XSLFTextShape) t)
                                .filter(r->r.getText().contains("EVE_")||r.getText().contains("REU_"))
                                .flatMap(p->Observable
                                        .just(p.getText().split("\\."))
                                        .map(r->{
                                            if(r[0].contains("EVE_"))
                                                existEvents=true;
                                            if (r[0].contains("REU_"))
                                                existReutilizable=true;
                                            addEventReutilizable(r);
                                            return r;
                                        })
                                )
                                .subscribe()
                        )
                        .subscribe()
                )
                .subscribe();

    }

    private static void addEventReutilizable(String[]r){

        listEventReu.add(new EventReutilizable()
                .code(r[0].trim())
                .name(r[1].trim()));
    }

    private static void addOtherEventsAndReutilizable(){

        if(existEvents){
            listEventReu.add(new EventReutilizable()
                    .code("EVE_NOT")
                    .name("Notificaciones"));
        }
        if(existReutilizable){
            listEventReu.add(new EventReutilizable()
                    .code("REU_NOT")
                    .name("Notificaciones"));
        }

    }

    private static void filter() {
        String processSolicitud = process+"_SIN_Solicitud";

        Flux.fromIterable(listRowData)
                .filter(t-> filters.stream()
                        .noneMatch(l -> l
                                .contains(t.getPlantilla())))
                .filter(r -> listEventReu
                        .stream()
                        .anyMatch(l -> l.code()
                                .contains(r.getProcReuEve()))
                ).
                filter(r->listVariEvit
                        .stream()
                        .noneMatch(l -> l
                                .equals(r.getVariable())))
                .filter(o->!processSolicitud.equals(o.getPlantilla()))
                .map(o -> {
                    RowData data = new RowData();
                    data.setFamily(o.getFamily());
                    data.setProcReuEve(o.getProcReuEve());
                    data.setPlantilla(o.getPlantilla());
                    data.setVariable(o.getVariable());
                    data.setJooScript(o.getJooScript());
                    listRowDataFilter.add(data);
                    return o;
                })
                //.groupBy(RowData::getVariable)
                //.flatMap(Flux::collectList)
                .subscribe();




    }

    private static void order()
    {

        Map<String, List<RowData>> noOfStocksByName = listRowDataFilter.stream()
                .collect(Collectors.groupingBy(RowData::getVariable));

        System.out.println(noOfStocksByName);



    }

}
