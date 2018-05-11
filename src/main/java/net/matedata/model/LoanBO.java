package net.matedata.model;

import java.math.BigDecimal;
import java.util.List;

public class LoanBO
{
    private BigDecimal loanAmountMax;
    private String loanEarliestDate;
    private String loanLatestDate;
    private Integer loanPlatformTotal;
    private Integer loanPlatformToday;
    private Integer loanPlatform3Days;
    private Integer loanPlatform7Days;
    private Integer loanPlatform15Days;
    private Integer loanPlatform30Days;
    private Integer loanPlatform60Days;
    private Integer loanPlatform90Days;
    private List<LoanPlatformBO> loanPlatforms;

    public BigDecimal getLoanAmountMax()
    {
        return loanAmountMax;
    }

    public void setLoanAmountMax(BigDecimal loanAmountMax)
    {
        this.loanAmountMax = loanAmountMax;
    }

    public String getLoanEarliestDate()
    {
        return loanEarliestDate;
    }

    public void setLoanEarliestDate(String loanEarliestDate)
    {
        this.loanEarliestDate = loanEarliestDate;
    }

    public String getLoanLatestDate()
    {
        return loanLatestDate;
    }

    public void setLoanLatestDate(String loanLatestDate)
    {
        this.loanLatestDate = loanLatestDate;
    }

    public Integer getLoanPlatformTotal()
    {
        return loanPlatformTotal;
    }

    public void setLoanPlatformTotal(Integer loanPlatformTotal)
    {
        this.loanPlatformTotal = loanPlatformTotal;
    }

    public Integer getLoanPlatformToday()
    {
        return loanPlatformToday;
    }

    public void setLoanPlatformToday(Integer loanPlatformToday)
    {
        this.loanPlatformToday = loanPlatformToday;
    }

    public Integer getLoanPlatform3Days()
    {
        return loanPlatform3Days;
    }

    public void setLoanPlatform3Days(Integer loanPlatform3Days)
    {
        this.loanPlatform3Days = loanPlatform3Days;
    }

    public Integer getLoanPlatform7Days()
    {
        return loanPlatform7Days;
    }

    public void setLoanPlatform7Days(Integer loanPlatform7Days)
    {
        this.loanPlatform7Days = loanPlatform7Days;
    }

    public Integer getLoanPlatform15Days()
    {
        return loanPlatform15Days;
    }

    public void setLoanPlatform15Days(Integer loanPlatform15Days)
    {
        this.loanPlatform15Days = loanPlatform15Days;
    }

    public Integer getLoanPlatform30Days()
    {
        return loanPlatform30Days;
    }

    public void setLoanPlatform30Days(Integer loanPlatform30Days)
    {
        this.loanPlatform30Days = loanPlatform30Days;
    }

    public Integer getLoanPlatform60Days()
    {
        return loanPlatform60Days;
    }

    public void setLoanPlatform60Days(Integer loanPlatform60Days)
    {
        this.loanPlatform60Days = loanPlatform60Days;
    }

    public Integer getLoanPlatform90Days()
    {
        return loanPlatform90Days;
    }

    public void setLoanPlatform90Days(Integer loanPlatform90Days)
    {
        this.loanPlatform90Days = loanPlatform90Days;
    }

    public List<LoanPlatformBO> getLoanPlatforms()
    {
        return loanPlatforms;
    }

    public void setLoanPlatforms(List<LoanPlatformBO> loanPlatforms)
    {
        this.loanPlatforms = loanPlatforms;
    }
}
