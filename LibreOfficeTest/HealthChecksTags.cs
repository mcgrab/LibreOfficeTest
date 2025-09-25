namespace LibreOfficeTest
{
    public static class HealthChecksTags
    {
        public static readonly string LivenessTag = "liveness";

        public static readonly string ReadinessTag = "readiness";

        public static readonly string[] LivenessTags = new[] { LivenessTag };

        public static readonly string[] ReadinessTags = new[] { ReadinessTag };
    }
}
