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
    private String yearMonth;
    private String dayStr;
    private String dayNic;
    private String dayFormat;
    private String select;
    private String amountStr;
    private BigDecimal amountNum;
    private String status;
    private Timestamp timestamp;

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

    public String getYearMonth() {
        return yearMonth;
    }

    public void setYearMonth(String yearMonth) {
        this.yearMonth = yearMonth;
    }

    public String getDayStr() {
        return dayStr;
    }

    public void setDayStr(String dayStr) {
        this.dayStr = dayStr;
    }

    public String getDayNic() {
        return dayNic;
    }

    public void setDayNic(String dayNic) {
        this.dayNic = dayNic;
    }

    public String getDayFormat() {
        return dayFormat;
    }

    public void setDayFormat(String dayFormat) {
        this.dayFormat = dayFormat;
    }

    public String getSelect() {
        return select;
    }

    public void setSelect(String select) {
        this.select = select;
    }

    public String getAmountStr() {
        return amountStr;
    }

    public void setAmountStr(String amountStr) {
        this.amountStr = amountStr;
    }

    public BigDecimal getAmountNum() {
        return amountNum;
    }

    public void setAmountNum(BigDecimal amountNum) {
        this.amountNum = amountNum;
    }

    public String getStatus() {
        return status;
    }

    public void setStatus(String status) {
        this.status = status;
    }

    public Timestamp getTimestamp() {
        return timestamp;
    }

    public void setTimestamp(Timestamp timestamp) {
        this.timestamp = timestamp;
    }
}
