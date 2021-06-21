import com.poiji.annotation.ExcelCell;
import com.poiji.annotation.ExcelSheet;
import lombok.*;
import lombok.experimental.Accessors;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@ToString
@ExcelSheet("RAW_DATA")
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

}
