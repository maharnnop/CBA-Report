﻿namespace BestPolicyReport.Models.OutputVatReport
{
    public class OutputVatReportInput
    {
        public string? InsurerCode { get; set; }
        public string? StartRpRefDate { get; set; }
        public string? EndRpRefDate { get; set; }
        public int? Limit { get; set; }
        public int? Pagecount { get; set; }
    }
}
