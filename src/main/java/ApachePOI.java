
import com.poiji.bind.Poiji;
import com.poiji.option.PoijiOptions;
import io.reactivex.Observable;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import reactor.core.publisher.Flux;

import java.io.*;
import java.util.ArrayList;
import java.util.List;


public class ApachePOI {

    public static List<EventReutilizable> listEventReu = new ArrayList<>();
    public static boolean existEvents=false;
    public static boolean existReutilizable=false;
    public  static List<RowData> rowData;

    public static void main(String...arg ) throws IOException {

        listEventReu.add(new EventReutilizable()
                .code("HAC_ACJ")
                .name("Aprobación cuentas justificativas de subvenciones"));

        PoijiOptions options = PoijiOptions.PoijiOptionsBuilder.settings().build();

        rowData = Poiji.fromExcel(new File("informe_plantillas.xlsx"), RowData.class, options);

        eventReutilizable();
        addOtherEventsAndReutilizable();
        filter();

    }

    private static void  eventReutilizable() throws IOException {

        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("HAC_ACJ.pptx"));
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

        Flux.fromIterable(rowData)
                .filter(r -> listEventReu
                        .stream()
                        .anyMatch(l -> l.code()
                                .contains(r.getProcReuEve())))
                .groupBy(RowData::getVariable)
                .flatMap(Flux::collectList)
                .subscribe(p->System.out.println("Data: "+p.toString()));


    }




}
