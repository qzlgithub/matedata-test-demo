package net.matedata.model;

public class LoanPlatformBO
{
    private String platformCode;
    private String platformType;
    private String loanEarliestDate;
    private String loanLatestDate;

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
}
