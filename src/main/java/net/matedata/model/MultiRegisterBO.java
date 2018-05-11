package net.matedata.model;

import java.util.List;

public class MultiRegisterBO
{
    private int registerInOneDayMax;
    private String registerEarliestDate;
    private String registerLatestDate;
    private int registerPlatformTotal;
    private int registerPlatformToday;
    private int registerPlatform3Days;
    private int registerPlatform7Days;
    private int registerPlatform15Days;
    private int registerPlatform30Days;
    private int registerPlatform60Days;
    private int registerPlatform90Days;
    private List<MultiRegisterPlatformBO> registerPlatforms;

    public int getRegisterInOneDayMax()
    {
        return registerInOneDayMax;
    }

    public void setRegisterInOneDayMax(int registerInOneDayMax)
    {
        this.registerInOneDayMax = registerInOneDayMax;
    }

    public String getRegisterEarliestDate()
    {
        return registerEarliestDate;
    }

    public void setRegisterEarliestDate(String registerEarliestDate)
    {
        this.registerEarliestDate = registerEarliestDate;
    }

    public String getRegisterLatestDate()
    {
        return registerLatestDate;
    }

    public void setRegisterLatestDate(String registerLatestDate)
    {
        this.registerLatestDate = registerLatestDate;
    }

    public int getRegisterPlatformTotal()
    {
        return registerPlatformTotal;
    }

    public void setRegisterPlatformTotal(int registerPlatformTotal)
    {
        this.registerPlatformTotal = registerPlatformTotal;
    }

    public int getRegisterPlatformToday()
    {
        return registerPlatformToday;
    }

    public void setRegisterPlatformToday(int registerPlatformToday)
    {
        this.registerPlatformToday = registerPlatformToday;
    }

    public int getRegisterPlatform3Days()
    {
        return registerPlatform3Days;
    }

    public void setRegisterPlatform3Days(int registerPlatform3Days)
    {
        this.registerPlatform3Days = registerPlatform3Days;
    }

    public int getRegisterPlatform7Days()
    {
        return registerPlatform7Days;
    }

    public void setRegisterPlatform7Days(int registerPlatform7Days)
    {
        this.registerPlatform7Days = registerPlatform7Days;
    }

    public int getRegisterPlatform15Days()
    {
        return registerPlatform15Days;
    }

    public void setRegisterPlatform15Days(int registerPlatform15Days)
    {
        this.registerPlatform15Days = registerPlatform15Days;
    }

    public int getRegisterPlatform30Days()
    {
        return registerPlatform30Days;
    }

    public void setRegisterPlatform30Days(int registerPlatform30Days)
    {
        this.registerPlatform30Days = registerPlatform30Days;
    }

    public int getRegisterPlatform60Days()
    {
        return registerPlatform60Days;
    }

    public void setRegisterPlatform60Days(int registerPlatform60Days)
    {
        this.registerPlatform60Days = registerPlatform60Days;
    }

    public int getRegisterPlatform90Days()
    {
        return registerPlatform90Days;
    }

    public void setRegisterPlatform90Days(int registerPlatform90Days)
    {
        this.registerPlatform90Days = registerPlatform90Days;
    }

    public List<MultiRegisterPlatformBO> getRegisterPlatforms()
    {
        return registerPlatforms;
    }

    public void setRegisterPlatforms(List<MultiRegisterPlatformBO> registerPlatforms)
    {
        this.registerPlatforms = registerPlatforms;
    }
}
