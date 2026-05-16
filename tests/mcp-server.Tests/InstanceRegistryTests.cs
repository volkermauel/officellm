using OfficeMcpServer.Models;

namespace OfficeMcpServer.Tests;

public class InstanceRegistryTests
{
    [Fact]
    public void RegisterInstance_ReturnsIncrementingIds()
    {
        var registry = new InstanceRegistry();
        var id1 = registry.RegisterInstance("PowerPoint", "deck1.pptx");
        var id2 = registry.RegisterInstance("PowerPoint", "deck2.pptx");

        Assert.Equal("powerpoint_1", id1);
        Assert.Equal("powerpoint_2", id2);
    }

    [Fact]
    public void RegisterInstance_AppearsInActiveInstances()
    {
        var registry = new InstanceRegistry();
        registry.RegisterInstance("PowerPoint", "test.pptx");

        var active = registry.GetActiveInstances();
        Assert.Single(active);
        Assert.Equal("PowerPoint", active[0].AppName);
        Assert.Equal("test.pptx", active[0].DocumentName);
    }

    [Fact]
    public void GetInstance_ReturnsInstance_WhenAlive()
    {
        var registry = new InstanceRegistry();
        var id = registry.RegisterInstance("PowerPoint", "deck.pptx");

        var instance = registry.GetInstance(id);
        Assert.NotNull(instance);
        Assert.Equal(id, instance.InstanceId);
    }

    [Fact]
    public void GetInstance_ReturnsNull_WhenNotFound()
    {
        var registry = new InstanceRegistry();
        var instance = registry.GetInstance("nonexistent");
        Assert.Null(instance);
    }

    [Fact]
    public void UpdateHeartbeat_UpdatesAppNameAndDocument()
    {
        var registry = new InstanceRegistry();
        var id = registry.RegisterInstance("PowerPoint", "old.pptx");

        registry.UpdateHeartbeat(id, "PowerPoint", "new.pptx");

        var instance = registry.GetInstance(id);
        Assert.NotNull(instance);
        Assert.Equal("new.pptx", instance.DocumentName);
    }

    [Fact]
    public void UpdateHeartbeat_DoesNothing_ForUnknownInstance()
    {
        var registry = new InstanceRegistry();
        // Should not throw
        registry.UpdateHeartbeat("nonexistent", "PowerPoint", "test.pptx");
    }

    [Fact]
    public void CleanupTimedOut_MarksStaleInstancesDead()
    {
        var registry = new InstanceRegistry();
        var id = registry.RegisterInstance("PowerPoint", "stale.pptx");

        // Manually expire the heartbeat by accessing the instance
        var instance = registry.GetInstance(id);
        Assert.NotNull(instance);
        // Simulate 60s passing by setting LastHeartbeat to the past
        instance.LastHeartbeat = DateTime.UtcNow.AddSeconds(-120);

        registry.CleanupTimedOut();

        // The instance should now be timed out and not returned
        var result = registry.GetInstance(id);
        Assert.Null(result);
    }

    [Fact]
    public void MultipleInstances_RegisterIndependently()
    {
        var registry = new InstanceRegistry();
        var id1 = registry.RegisterInstance("PowerPoint", "a.pptx");
        var id2 = registry.RegisterInstance("PowerPoint", "b.pptx");

        var active = registry.GetActiveInstances();
        Assert.Equal(2, active.Count);
        Assert.Contains(active, i => i.InstanceId == id1);
        Assert.Contains(active, i => i.InstanceId == id2);
    }

    [Fact]
    public void ActiveInstances_ExcludesTimedOut()
    {
        var registry = new InstanceRegistry();
        var id1 = registry.RegisterInstance("PowerPoint", "alive.pptx");
        var id2 = registry.RegisterInstance("PowerPoint", "dead.pptx");

        // Expire the second instance
        var dead = registry.GetInstance(id2);
        Assert.NotNull(dead);
        dead.LastHeartbeat = DateTime.UtcNow.AddSeconds(-120);

        var active = registry.GetActiveInstances();
        Assert.Single(active);
        Assert.Equal(id1, active[0].InstanceId);
    }

    [Fact]
    public void RegisterInstance_HostSpecificPrefixes()
    {
        var registry = new InstanceRegistry();
        var ppt = registry.RegisterInstance("PowerPoint", "deck.pptx");
        var word = registry.RegisterInstance("Word", "doc.docx");
        var excel = registry.RegisterInstance("Excel", "sheet.xlsx");
        var outlook = registry.RegisterInstance("Outlook", "inbox");
        var unknown = registry.RegisterInstance("CustomApp", "test");

        Assert.Equal("powerpoint_1", ppt);
        Assert.Equal("word_2", word);
        Assert.Equal("excel_3", excel);
        Assert.Equal("outlook_4", outlook);
        Assert.Equal("office_5", unknown);
    }

    [Fact]
    public void RegisterInstance_MixedHosts_IncrementCorrectly()
    {
        var registry = new InstanceRegistry();
        var id1 = registry.RegisterInstance("PowerPoint", "a.pptx");
        var id2 = registry.RegisterInstance("Word", "b.docx");
        var id3 = registry.RegisterInstance("PowerPoint", "c.pptx");

        Assert.Equal("powerpoint_1", id1);
        Assert.Equal("word_2", id2);
        Assert.Equal("powerpoint_3", id3);
    }
}
