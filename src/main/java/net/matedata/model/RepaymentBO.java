package net.matedata.model;

import java.math.BigDecimal;
import java.util.List;

public class RepaymentBO
{
    private BigDecimal repaymentAmountMax;
    private String repaymentEarliestDate;
    private String repaymentLatestDate;
    private Integer repaymentPlatformTotal;
    private Integer repaymentPlatformToday;
    private Integer repaymentPlatform3Days;
    private Integer repaymentPlatform7Days;
    private Integer repaymentPlatform15Days;
    private Integer repaymentPlatform30Days;
    private Integer repaymentPlatform60Days;
    private Integer repaymentPlatform90Days;
    private List<RepaymentPlatformBO> repaymentPlatforms;

    public BigDecimal getRepaymentAmountMax()
    {
        return repaymentAmountMax;
    }

    public void setRepaymentAmountMax(BigDecimal repaymentAmountMax)
    {
        this.repaymentAmountMax = repaymentAmountMax;
    }

    public String getRepaymentEarliestDate()
    {
        return repaymentEarliestDate;
    }

    public void setRepaymentEarliestDate(String repaymentEarliestDate)
    {
        this.repaymentEarliestDate = repaymentEarliestDate;
    }

    public String getRepaymentLatestDate()
    {
        return repaymentLatestDate;
    }

    public void setRepaymentLatestDate(String repaymentLatestDate)
    {
        this.repaymentLatestDate = repaymentLatestDate;
    }

    public Integer getRepaymentPlatformTotal()
    {
        return repaymentPlatformTotal;
    }

    public void setRepaymentPlatformTotal(Integer repaymentPlatformTotal)
    {
        this.repaymentPlatformTotal = repaymentPlatformTotal;
    }

    public Integer getRepaymentPlatformToday()
    {
        return repaymentPlatformToday;
    }

    public void setRepaymentPlatformToday(Integer repaymentPlatformToday)
    {
        this.repaymentPlatformToday = repaymentPlatformToday;
    }

    public Integer getRepaymentPlatform3Days()
    {
        return repaymentPlatform3Days;
    }

    public void setRepaymentPlatform3Days(Integer repaymentPlatform3Days)
    {
        this.repaymentPlatform3Days = repaymentPlatform3Days;
    }

    public Integer getRepaymentPlatform7Days()
    {
        return repaymentPlatform7Days;
    }

    public void setRepaymentPlatform7Days(Integer repaymentPlatform7Days)
    {
        this.repaymentPlatform7Days = repaymentPlatform7Days;
    }

    public Integer getRepaymentPlatform15Days()
    {
        return repaymentPlatform15Days;
    }

    public void setRepaymentPlatform15Days(Integer repaymentPlatform15Days)
    {
        this.repaymentPlatform15Days = repaymentPlatform15Days;
    }

    public Integer getRepaymentPlatform30Days()
    {
        return repaymentPlatform30Days;
    }

    public void setRepaymentPlatform30Days(Integer repaymentPlatform30Days)
    {
        this.repaymentPlatform30Days = repaymentPlatform30Days;
    }

    public Integer getRepaymentPlatform60Days()
    {
        return repaymentPlatform60Days;
    }

    public void setRepaymentPlatform60Days(Integer repaymentPlatform60Days)
    {
        this.repaymentPlatform60Days = repaymentPlatform60Days;
    }

    public Integer getRepaymentPlatform90Days()
    {
        return repaymentPlatform90Days;
    }

    public void setRepaymentPlatform90Days(Integer repaymentPlatform90Days)
    {
        this.repaymentPlatform90Days = repaymentPlatform90Days;
    }

    public List<RepaymentPlatformBO> getRepaymentPlatforms()
    {
        return repaymentPlatforms;
    }

    public void setRepaymentPlatforms(List<RepaymentPlatformBO> repaymentPlatforms)
    {
        this.repaymentPlatforms = repaymentPlatforms;
    }
}
