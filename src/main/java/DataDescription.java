import com.poiji.annotation.ExcelCell;
import com.poiji.annotation.ExcelSheet;
import lombok.*;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@ToString
@ExcelSheet("Diccionario ")
public class DataDescription {

    @ExcelCell(0)
    private String codigo;

    @ExcelCell(1)
    private String descripcion;

    @ExcelCell(3)
    private String tipo;

    @ExcelCell(4)
    private String categorizacion;

}
