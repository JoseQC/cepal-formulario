import com.poiji.annotation.ExcelCell;
import com.poiji.annotation.ExcelSheet;
import lombok.*;
import lombok.experimental.Accessors;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
@ToString
@ExcelSheet("Hoja1")
@Accessors(fluent = true)
public class VariableEvitar {
    @ExcelCell(0)
    private String variable;
}
