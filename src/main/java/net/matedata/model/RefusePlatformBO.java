package net.matedata.model;

public class RefusePlatformBO
{
    private String platformCode;
    private String platformType;
    private String refuseEarliestDate;
    private String refuseLatestDate;

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
}
