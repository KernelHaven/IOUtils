package net.ssehub.kernel_haven.io.csv;

import com.opencsv.bean.CsvBindByName;

@Deprecated
public class VariableValueBean {
    
    @CsvBindByName(column = "Name", required = true)
    private String name;

    @CsvBindByName(column = "Value", required = true)
    private String value;

    /**
     * @return the name
     */
    String getName() {
        return name;
    }

    /**
     * @param name the name to set
     */
    void setName(String name) {
        this.name = name;
    }

    /**
     * @return the value
     */
    String getValue() {
        return value;
    }

    /**
     * @param value the value to set
     */
    void setValue(String value) {
        this.value = value;
    }

}
