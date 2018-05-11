package net.matedata.model;

public class RepaymentPlatformBO
{
    private String platformCode;
    private String platformType;
    private String repaymentEarliestDate;
    private String repaymentLatestDate;

    public String getPlatformCode()
    {
        return platformCode;
    }

    public void setPlatformCode(String platformCode)
    {
        this.platformCode = platformCode;
    }

    public String getPlatformType()
    {
        return platformType;
    }

    public void setPlatformType(String platformType)
    {
        this.platformType = platformType;
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
}
