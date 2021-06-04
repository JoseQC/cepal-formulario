import lombok.Getter;
import lombok.Setter;
import lombok.experimental.Accessors;

@Getter
@Setter@Accessors(fluent = true)
public class RowData {

    private String family;
    private String typeProReuEve;
    private String template;
    private String variable;
    private Boolean jooScript;

    public String getFamily() {
        return family;
    }

    public String getTypeProReuEve() {
        return typeProReuEve;
    }

    public String getTemplate() {
        return template;
    }

    public String getVariable() {
        return variable;
    }

    public Boolean getJooScript() {
        return jooScript;
    }

    @Override
    public String toString() {
        return "RowData{" +
                "family='" + family + '\'' +
                ", typeProReuEve='" + typeProReuEve + '\'' +
                ", template='" + template + '\'' +
                ", variable='" + variable + '\'' +
                ", jooScript=" + jooScript +
                '}';
    }
}
