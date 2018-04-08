package excel;

public class Execls {
    private String x;
    private String y;
    private String value;

    public String getX() {
        return x;
    }

    public void setX(String x) {
        this.x = x;
    }

    public String getY() {
        return y;
    }

    public void setY(String y) {
        this.y = y;
    }

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public Execls() {

    }

    public Execls(String x, String y, String value) {
        this.x = x;
        this.y = y;
        this.value = value;
    }
}
