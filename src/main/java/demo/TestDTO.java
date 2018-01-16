package demo;

import java.math.BigDecimal;
import java.sql.Timestamp;
import java.util.Date;

public class TestDTO {

    private String name;
    private String idCardOne;
    private String idCardTwo;
    private String phone;
    private Date birthday;
    private String yearMonthStr;
    private Integer yearMonth;
    private Integer dayStr;
    private String dayNic;
    private Long dayFormat;
    private String select;
    private Double amountStr;
    private BigDecimal amountNum;
    private Integer status;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getIdCardOne() {
        return idCardOne;
    }

    public void setIdCardOne(String idCardOne) {
        this.idCardOne = idCardOne;
    }

    public String getIdCardTwo() {
        return idCardTwo;
    }

    public void setIdCardTwo(String idCardTwo) {
        this.idCardTwo = idCardTwo;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public Date getBirthday() {
        return birthday;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }

    public String getYearMonthStr() {
        return yearMonthStr;
    }

    public void setYearMonthStr(String yearMonthStr) {
        this.yearMonthStr = yearMonthStr;
    }

    public Integer getYearMonth() {
        return yearMonth;
    }

    public void setYearMonth(Integer yearMonth) {
        this.yearMonth = yearMonth;
    }

    public Integer getDayStr() {
        return dayStr;
    }

    public void setDayStr(Integer dayStr) {
        this.dayStr = dayStr;
    }

    public String getDayNic() {
        return dayNic;
    }

    public void setDayNic(String dayNic) {
        this.dayNic = dayNic;
    }

    public Long getDayFormat() {
        return dayFormat;
    }

    public void setDayFormat(Long dayFormat) {
        this.dayFormat = dayFormat;
    }

    public String getSelect() {
        return select;
    }

    public void setSelect(String select) {
        this.select = select;
    }

    public Double getAmountStr() {
        return amountStr;
    }

    public void setAmountStr(Double amountStr) {
        this.amountStr = amountStr;
    }

    public BigDecimal getAmountNum() {
        return amountNum;
    }

    public void setAmountNum(BigDecimal amountNum) {
        this.amountNum = amountNum;
    }

    public Integer getStatus() {
        return status;
    }

    public void setStatus(Integer status) {
        this.status = status;
    }
}
