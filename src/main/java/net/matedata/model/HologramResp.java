package net.matedata.model;

import com.alibaba.fastjson.JSONObject;

public class HologramResp
{
    private Person person;
    private Integer hit;
    private Integer blacklistCode;
    private Integer overdueCode;
    private JSONObject overdueRes;
    private Integer multiRegisterCode;
    private JSONObject multiRegisterRes;
    private Integer loanCode;
    private JSONObject loanRes;
    private Integer refuseCode;
    private JSONObject refuseRes;
    private Integer repaymentCode;
    private JSONObject repaymentRes;

    public Person getPerson()
    {
        return person;
    }

    public void setPerson(Person person)
    {
        this.person = person;
    }

    public Integer getHit()
    {
        return hit;
    }

    public void setHit(Integer hit)
    {
        this.hit = hit;
    }

    public Integer getBlacklistCode()
    {
        return blacklistCode;
    }

    public void setBlacklistCode(Integer blacklistCode)
    {
        this.blacklistCode = blacklistCode;
    }

    public Integer getOverdueCode()
    {
        return overdueCode;
    }

    public void setOverdueCode(Integer overdueCode)
    {
        this.overdueCode = overdueCode;
    }

    public JSONObject getOverdueRes()
    {
        return overdueRes;
    }

    public void setOverdueRes(JSONObject overdueRes)
    {
        this.overdueRes = overdueRes;
    }

    public Integer getMultiRegisterCode()
    {
        return multiRegisterCode;
    }

    public void setMultiRegisterCode(Integer multiRegisterCode)
    {
        this.multiRegisterCode = multiRegisterCode;
    }

    public JSONObject getMultiRegisterRes()
    {
        return multiRegisterRes;
    }

    public void setMultiRegisterRes(JSONObject multiRegisterRes)
    {
        this.multiRegisterRes = multiRegisterRes;
    }

    public Integer getLoanCode()
    {
        return loanCode;
    }

    public void setLoanCode(Integer loanCode)
    {
        this.loanCode = loanCode;
    }

    public JSONObject getLoanRes()
    {
        return loanRes;
    }

    public void setLoanRes(JSONObject loanRes)
    {
        this.loanRes = loanRes;
    }

    public Integer getRefuseCode()
    {
        return refuseCode;
    }

    public void setRefuseCode(Integer refuseCode)
    {
        this.refuseCode = refuseCode;
    }

    public JSONObject getRefuseRes()
    {
        return refuseRes;
    }

    public void setRefuseRes(JSONObject refuseRes)
    {
        this.refuseRes = refuseRes;
    }

    public Integer getRepaymentCode()
    {
        return repaymentCode;
    }

    public void setRepaymentCode(Integer repaymentCode)
    {
        this.repaymentCode = repaymentCode;
    }

    public JSONObject getRepaymentRes()
    {
        return repaymentRes;
    }

    public void setRepaymentRes(JSONObject repaymentRes)
    {
        this.repaymentRes = repaymentRes;
    }
}
