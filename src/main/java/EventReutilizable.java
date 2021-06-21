import lombok.*;
import lombok.experimental.Accessors;

import java.util.List;

@Getter
@AllArgsConstructor
@NoArgsConstructor
@ToString
@Setter()
@Accessors(fluent = true)
public class EventReutilizable {
    private String code;
    private String name;
    private List<String> variable;
}
