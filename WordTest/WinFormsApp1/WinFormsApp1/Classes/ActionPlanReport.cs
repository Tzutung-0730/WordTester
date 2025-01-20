namespace WordTester.Classes;

internal class ActionPlanReport
{
    public int IssueSeqNo { get; set; }
    public string IssueName { get; set; }
    public int PlanSeqNo { get; set; }
    public string PlanName { get; set; }
    public string StakeHolders { get; set; }
    public int SeasonSeqNo { get; set; }
    public string SectionNo { get; set; }
    public string ActionPlan { get; set; }
	public string EstimatedDelivery { get; set; }
    public string ActualDelivery { get; set; }
    public string OnTrackDesc { get; set; }
    public string OffTrackDesc { get; set; }
}
