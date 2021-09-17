import com.poiji.annotation.ExcelCell;
import com.poiji.annotation.ExcelCellName;
import com.poiji.annotation.ExcelSheet;
import lombok.*;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@ToString
@ExcelSheet("Hoja2")
public class RowDataDTO {

    @ExcelCell(0)
    @ExcelCellName("Familia")
    private String family;

    @ExcelCell(1)
    @ExcelCellName("Proc / Reu / Eve")
    private String procReuEve;

    @ExcelCell(2)
    @ExcelCellName("Plantilla")
    private String plantilla;

    @ExcelCell(3)
    @ExcelCellName("Variable")
    private String variable;

    @ExcelCell(4)
    @ExcelCellName("JooScript")
    private String jooScript;

    @ExcelCellName("Orden")
    private Integer order;
}
