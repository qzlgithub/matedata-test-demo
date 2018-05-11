package net.matedata.model;

public class OverduePlatformBO
{
    private String platformCode;
    private String platformType;
    private String overdueEarliestTime;
    private String overdueLatestTime;

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

    public String getOverdueEarliestTime()
    {
        return overdueEarliestTime;
    }

    public void setOverdueEarliestTime(String overdueEarliestTime)
    {
        this.overdueEarliestTime = overdueEarliestTime;
    }

    public String getOverdueLatestTime()
    {
        return overdueLatestTime;
    }

    public void setOverdueLatestTime(String overdueLatestTime)
    {
        this.overdueLatestTime = overdueLatestTime;
    }
}
