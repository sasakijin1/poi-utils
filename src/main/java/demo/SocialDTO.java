package demo;

import java.math.BigDecimal;
import java.util.Date;

public class SocialDTO {
    private String empName;

    private String companyId;

    private Integer socialRuleId;

    private Integer fundRuleId;

    private Integer idCardType;

    private String idCard;

    private Date inDate;

    private String phone;

    private Integer status;

    private BigDecimal companyBase;

    private BigDecimal companyRatio;

    private BigDecimal companyAmount;

    private BigDecimal personalBase;

    private BigDecimal personalRatio;

    private BigDecimal personalAmount;

    public String getEmpName() {
        return empName;
    }

    public void setEmpName(String empName) {
        this.empName = empName;
    }

    public String getCompanyId() {
        return companyId;
    }

    public void setCompanyId(String companyId) {
        this.companyId = companyId;
    }

    public Integer getSocialRuleId() {
        return socialRuleId;
    }

    public void setSocialRuleId(Integer socialRuleId) {
        this.socialRuleId = socialRuleId;
    }

    public Integer getFundRuleId() {
        return fundRuleId;
    }

    public void setFundRuleId(Integer fundRuleId) {
        this.fundRuleId = fundRuleId;
    }

    public Integer getIdCardType() {
        return idCardType;
    }

    public void setIdCardType(Integer idCardType) {
        this.idCardType = idCardType;
    }

    public String getIdCard() {
        return idCard;
    }

    public void setIdCard(String idCard) {
        this.idCard = idCard;
    }

    public Date getInDate() {
        return inDate;
    }

    public void setInDate(Date inDate) {
        this.inDate = inDate;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    public Integer getStatus() {
        return status;
    }

    public void setStatus(Integer status) {
        this.status = status;
    }

    public BigDecimal getCompanyBase() {
        return companyBase;
    }

    public void setCompanyBase(BigDecimal companyBase) {
        this.companyBase = companyBase;
    }

    public BigDecimal getCompanyRatio() {
        return companyRatio;
    }

    public void setCompanyRatio(BigDecimal companyRatio) {
        this.companyRatio = companyRatio;
    }

    public BigDecimal getCompanyAmount() {
        return companyAmount;
    }

    public void setCompanyAmount(BigDecimal companyAmount) {
        this.companyAmount = companyAmount;
    }

    public BigDecimal getPersonalBase() {
        return personalBase;
    }

    public void setPersonalBase(BigDecimal personalBase) {
        this.personalBase = personalBase;
    }

    public BigDecimal getPersonalRatio() {
        return personalRatio;
    }

    public void setPersonalRatio(BigDecimal personalRatio) {
        this.personalRatio = personalRatio;
    }

    public BigDecimal getPersonalAmount() {
        return personalAmount;
    }

    public void setPersonalAmount(BigDecimal personalAmount) {
        this.personalAmount = personalAmount;
    }
}
