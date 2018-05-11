package net.matedata.model;

import java.math.BigDecimal;
import java.util.List;

public class RefuseBO
{
    private BigDecimal refuseAmountMax;
    private String refuseEarliestDate;
    private String refuseLatestDate;
    private Integer refusePlatformTotal;
    private Integer refusePlatformToday;
    private Integer refusePlatform3Days;
    private Integer refusePlatform7Days;
    private Integer refusePlatform15Days;
    private Integer refusePlatform30Days;
    private Integer refusePlatform60Days;
    private Integer refusePlatform90Days;
    private List<RefusePlatformBO> refusePlatforms;

    public BigDecimal getRefuseAmountMax()
    {
        return refuseAmountMax;
    }

    public void setRefuseAmountMax(BigDecimal refuseAmountMax)
    {
        this.refuseAmountMax = refuseAmountMax;
    }

    public String getRefuseEarliestDate()
    {
        return refuseEarliestDate;
    }

    public void setRefuseEarliestDate(String refuseEarliestDate)
    {
        this.refuseEarliestDate = refuseEarliestDate;
    }

    public String getRefuseLatestDate()
    {
        return refuseLatestDate;
    }

    public void setRefuseLatestDate(String refuseLatestDate)
    {
        this.refuseLatestDate = refuseLatestDate;
    }

    public Integer getRefusePlatformTotal()
    {
        return refusePlatformTotal;
    }

    public void setRefusePlatformTotal(Integer refusePlatformTotal)
    {
        this.refusePlatformTotal = refusePlatformTotal;
    }

    public Integer getRefusePlatformToday()
    {
        return refusePlatformToday;
    }

    public void setRefusePlatformToday(Integer refusePlatformToday)
    {
        this.refusePlatformToday = refusePlatformToday;
    }

    public Integer getRefusePlatform3Days()
    {
        return refusePlatform3Days;
    }

    public void setRefusePlatform3Days(Integer refusePlatform3Days)
    {
        this.refusePlatform3Days = refusePlatform3Days;
    }

    public Integer getRefusePlatform7Days()
    {
        return refusePlatform7Days;
    }

    public void setRefusePlatform7Days(Integer refusePlatform7Days)
    {
        this.refusePlatform7Days = refusePlatform7Days;
    }

    public Integer getRefusePlatform15Days()
    {
        return refusePlatform15Days;
    }

    public void setRefusePlatform15Days(Integer refusePlatform15Days)
    {
        this.refusePlatform15Days = refusePlatform15Days;
    }

    public Integer getRefusePlatform30Days()
    {
        return refusePlatform30Days;
    }

    public void setRefusePlatform30Days(Integer refusePlatform30Days)
    {
        this.refusePlatform30Days = refusePlatform30Days;
    }

    public Integer getRefusePlatform60Days()
    {
        return refusePlatform60Days;
    }

    public void setRefusePlatform60Days(Integer refusePlatform60Days)
    {
        this.refusePlatform60Days = refusePlatform60Days;
    }

    public Integer getRefusePlatform90Days()
    {
        return refusePlatform90Days;
    }

    public void setRefusePlatform90Days(Integer refusePlatform90Days)
    {
        this.refusePlatform90Days = refusePlatform90Days;
    }

    public List<RefusePlatformBO> getRefusePlatforms()
    {
        return refusePlatforms;
    }

    public void setRefusePlatforms(List<RefusePlatformBO> refusePlatforms)
    {
        this.refusePlatforms = refusePlatforms;
    }
}
