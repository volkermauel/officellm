using OfficeMcpServer.Models;

namespace OfficeMcpServer.Tests;

public class AuditLogTests : IDisposable
{
    private readonly string _testDir;
    private readonly AuditLog _log;

    public AuditLogTests()
    {
        _testDir = Path.Combine(Path.GetTempPath(), $"audit-test-{Guid.NewGuid()}");
        Directory.CreateDirectory(_testDir);
        _log = new AuditLog(_testDir);
    }

    public void Dispose()
    {
        _log.Dispose();
        if (Directory.Exists(_testDir))
            Directory.Delete(_testDir, true);
    }

    [Fact]
    public void Log_AppendsEntry_ToFile()
    {
        var entry = new AuditEntry
        {
            ToolName = "powerpoint_get_deck_outline",
            InstanceId = "powerpoint_1",
            Outcome = "success"
        };

        _log.Log(entry);

        var lines = File.ReadAllLines(Path.Combine(_testDir, "audit.log"));
        Assert.Single(lines);

        var parsed = System.Text.Json.JsonSerializer.Deserialize<AuditEntry>(lines[0]);
        Assert.NotNull(parsed);
        Assert.Equal("powerpoint_get_deck_outline", parsed.ToolName);
        Assert.Equal("success", parsed.Outcome);
    }

    [Fact]
    public void Log_MultipleEntries_CreatesMultipleLines()
    {
        _log.Log(new AuditEntry { ToolName = "tool1", Outcome = "success" });
        _log.Log(new AuditEntry { ToolName = "tool2", Outcome = "error" });

        var lines = File.ReadAllLines(Path.Combine(_testDir, "audit.log"));
        Assert.Equal(2, lines.Length);
    }

    [Fact]
    public void Log_EntryContainsRequiredFields()
    {
        var entry = new AuditEntry
        {
            ToolName = "test",
            InstanceId = "ppt_1",
            Inputs = "{\"slideIndex\":0}",
            RequiresConfirmation = true,
            ConfirmationStatus = "pending",
            Outcome = "success"
        };

        _log.Log(entry);

        var lines = File.ReadAllLines(Path.Combine(_testDir, "audit.log"));
        var parsed = System.Text.Json.JsonSerializer.Deserialize<AuditEntry>(lines[0]);

        Assert.NotEmpty(parsed.Timestamp);
        Assert.Equal("test", parsed.ToolName);
        Assert.Equal("ppt_1", parsed.InstanceId);
        Assert.Equal("{\"slideIndex\":0}", parsed.Inputs);
        Assert.True(parsed.RequiresConfirmation);
        Assert.Equal("pending", parsed.ConfirmationStatus);
        Assert.NotEmpty(parsed.AuditId);
    }

    [Fact]
    public void Log_ThreadSafe_ConcurrentWrites()
    {
        var tasks = Enumerable.Range(0, 10).Select(i => Task.Run(() =>
        {
            _log.Log(new AuditEntry { ToolName = $"concurrent-{i}", Outcome = "ok" });
        }));

        Task.WaitAll(tasks.ToArray());

        var lines = File.ReadAllLines(Path.Combine(_testDir, "audit.log"));
        Assert.Equal(10, lines.Length);
    }

    [Fact]
    public void AuditEntry_GeneratesTimestampAndAuditId()
    {
        var entry = new AuditEntry();

        Assert.NotEmpty(entry.Timestamp);
        Assert.NotEmpty(entry.AuditId);
        Assert.True(entry.Timestamp.Contains("T"), "Timestamp should be ISO 8601");
    }
}
