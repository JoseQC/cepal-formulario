import com.poiji.annotation.ExcelCell;
import com.poiji.annotation.ExcelSheet;
import lombok.*;
import lombok.experimental.Accessors;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@ToString
@Accessors(fluent = true)
@ExcelSheet("RAW_DATA")
//@ExcelSheet("Hoja2")
public class RowData {

    @ExcelCell(0)
    private String family;

    @ExcelCell(1)
    private String procReuEve;

    @ExcelCell(2)
    private String plantilla;

    @ExcelCell(3)
    private String variable;

    @ExcelCell(4)
    private String jooScript;

    private Integer order;

}
