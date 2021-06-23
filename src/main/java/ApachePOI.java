
import com.poiji.bind.Poiji;
import com.poiji.option.PoijiOptions;
import io.reactivex.Observable;
import io.reactivex.disposables.Disposable;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import reactor.core.publisher.Flux;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;


public class ApachePOI {

    public static List<EventReutilizable> listEventReu = new ArrayList<>();
    public static List<VariableEvitar>listVariEvit= new ArrayList<>();
    public static boolean existEvents=false;
    public static boolean existReutilizable=false;
    public  static List<RowData> listRowData;
    public static List<RowData> listFilter = new ArrayList<>(); 
    public  static List<DataDescription> listDataDescription;
    public  static PoijiOptions options = PoijiOptions.PoijiOptionsBuilder.settings().build();

    public static void main(String...arg ) throws IOException {


        listEventReu.add(new EventReutilizable()
                .code("HAC_ACJ")
                .name("Aprobación cuentas justificativas de subvenciones"));

        //PoijiOptions options = PoijiOptions.PoijiOptionsBuilder.settings().build();
        listRowData = Poiji.fromExcel(new File("informe_plantillas.xlsx"), RowData.class, options);

        eventReutilizable();
        readVariableEvitar();
        addOtherEventsAndReutilizable();
        listEventReu.forEach(o->System.out.println(""+o.toString()));
        filter();
        extractDescription();
        filterAdvanced();
        System.out.println(listFilter.size());
       // readTemplates();

    }
    private static void readVariableEvitar(){
        listVariEvit = Poiji.fromExcel(new File("Variables a evitar.xlsx"), VariableEvitar.class, options);
       System.out.println(listVariEvit.size());
    }

    private static  void readTemplates() {

        //set path of file
        File file = new File("C:\\Users\\Jose\\Workspace - Developer\\Java\\Cep@l\\Cep@l - Formularios\\cepal-formulario\\7.2.Hacienda\\HAC_ACJ\\HAC_ACJ");
        String[] list = file.list();

        listRowData = Poiji.fromExcel(new File("informe_plantillas.xlsx"), RowData.class, options);

        //RxJava
        Disposable subscribe = Observable.fromIterable(listRowData).map(r -> {
            assert list != null;
            for (String data : list) {
                String[] arrOfStr = data.split("\\.");
                if (r.plantilla().endsWith(arrOfStr[1])) {
                    r.order(Integer.parseInt(arrOfStr[0]));
                }
            }
            return r;
        }).subscribe();

    }

    private static void extractDescription(){

        listDataDescription = Poiji.fromExcel(new File("Diccionario_variables.xlsx"), DataDescription.class, options);
        //System.out.print(listDataDescription.size());
    }

    private static void  eventReutilizable() throws IOException {

        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("C:\\Users\\Jose\\Workspace - Developer\\Java\\Cep@l\\Cep@l - Formularios\\cepal-formulario\\7.2.Hacienda\\HAC_ACJ\\HAC_ACJ.pptx"));
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

       Flux.fromIterable(listRowData)
                .filter(r -> listEventReu
                        .stream()
                        .anyMatch(l -> l.code()
                                .contains(r.procReuEve())))
                .filter(t->listVariEvit
                            .stream()
                            .noneMatch(y -> y.variable()
                                    .contains(t.variable()))
                )
               .map(y->{
                   listFilter.add(new RowData()
                           .family(y.family())
                           .procReuEve(y.procReuEve())
                           .plantilla(y.plantilla())
                           .variable(y.variable())
                           .jooScript(y.jooScript()));
                   return y;
               })
               .subscribe();


    }

    private static void filterAdvanced(){

        Map<String, List<RowData>> groupByVariable =
                listFilter.stream().collect(Collectors.groupingBy(RowData::variable));
        System.out.println(groupByVariable.size());



    }

}
