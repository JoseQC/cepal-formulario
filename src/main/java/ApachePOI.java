
import com.poiji.bind.Poiji;
import com.poiji.option.PoijiOptions;
import io.reactivex.Observable;
import io.reactivex.disposables.Disposable;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import reactor.core.publisher.Flux;

import java.io.*;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.Stream;


public class ApachePOI {

    public static List<EventReutilizable> listEventReu = new ArrayList<>();

    public static List<String> listVariEvit = new ArrayList<>();
    public static boolean existEvents=false;
    public static boolean existReutilizable=false;
    public  static List<RowData> listRowData ;
    public  static List<RowData> listRowDataFilter = new ArrayList<>();
    public  static List<DataDescription> listDataDescription;
    public  static PoijiOptions options = PoijiOptions.PoijiOptionsBuilder.settings().build();
    public static String process = "PER_CAR";

    public static void main(String...arg ) throws IOException {




        listEventReu.add(new EventReutilizable()
                .code(process)
                .name("Aprobación cuentas justificativas de subvenciones"));

        //PoijiOptions options = PoijiOptions.PoijiOptionsBuilder.settings().build();
        listRowData = Poiji.fromExcel(new File("C:\\Users\\jquipsec\\Documents\\cep@l\\informe_plantillas.xlsx"), RowData.class, options);

        variablesEvitar();
        eventReutilizable();
        listEventReu.forEach(o->System.out.println(""+o.toString()));
        filter();
        listRowDataFilter.forEach(System.out::println);
        order();


    }

    private static void variablesEvitar() throws IOException {


        File file = new File("C:\\Users\\jquipsec\\Documents\\Workspace - Jose\\Developer\\Java\\Cep@l\\appFilterVariables\\cepal-formulario\\Variables a evitar.xlsx");
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
        File file = new File("C:\\Users\\jquipsec\\Documents\\cep@l\\7.1.Contratación\\CON_CDC\\CON_CDC");
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

        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("C:\\Users\\jquipsec\\Documents\\cep@l\\7.4.Personal\\PER_CAR\\PER_CAR.pptx"));
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
        String processSolicitud = process+"_SIN_Solicitud_interesado";

        Flux.fromIterable(listRowData)
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
