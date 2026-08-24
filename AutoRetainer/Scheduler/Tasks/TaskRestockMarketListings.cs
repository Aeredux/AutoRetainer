using AutoRetainer.Scheduler.Handlers;
using AutoRetainerAPI.Configuration;
using ECommons.Automation;
using ECommons.ExcelServices;
using ECommons.GameHelpers;
using ECommons.Throttlers;
using ECommons.UIHelpers.AddonMasterImplementations;
using FFXIVClientStructs.FFXIV.Client.UI;
using FFXIVClientStructs.FFXIV.Client.Game.Event;
using FFXIVClientStructs.FFXIV.Client.Game;
using FFXIVClientStructs.FFXIV.Client.UI.Agent;
using FFXIVClientStructs.FFXIV.Component.GUI;
using Lumina.Excel.Sheets;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AutoRetainer.Scheduler.Tasks;

internal static unsafe class TaskRestockMarketListings
{
    internal readonly struct ExperimentalDebugState
    {
        public readonly bool InjectionFeatureCompiled;
        public readonly bool FollowUpClosePulseCompiled;
        public readonly bool GuardedConfirmFeatureCompiled;
        public readonly bool GuardedConfirmRuntimeEnabled;
        public readonly bool GuardedConfirmEffectiveEnabled;
        public readonly bool GuardedConfirmOneShotArmed;
        public readonly bool GuardedConfirmBlockedByPlugin;
        public readonly int PendingType136RetainerMarketOps;
        public readonly bool DisabledForSession;
        public readonly bool GuardedConfirmDisabledForSession;
        public readonly bool AwaitingAck;
        public readonly bool GuardedConfirmAwaitingAck;
        public readonly bool InjectedThisListing;
        public readonly bool GuardedConfirmClickedThisListing;
        public readonly uint ContextId;
        public readonly uint GuardedConfirmContextId;
        public readonly long AckTimeoutInMs;
        public readonly long GuardedConfirmAckTimeoutInMs;

        public ExperimentalDebugState(bool injectionFeatureCompiled, bool followUpClosePulseCompiled, bool guardedConfirmFeatureCompiled, bool guardedConfirmRuntimeEnabled, bool guardedConfirmEffectiveEnabled, bool guardedConfirmOneShotArmed, bool guardedConfirmBlockedByPlugin, int pendingType136RetainerMarketOps, bool disabledForSession, bool guardedConfirmDisabledForSession, bool awaitingAck, bool guardedConfirmAwaitingAck, bool injectedThisListing, bool guardedConfirmClickedThisListing, uint contextId, uint guardedConfirmContextId, long ackTimeoutInMs, long guardedConfirmAckTimeoutInMs)
        {
            InjectionFeatureCompiled = injectionFeatureCompiled;
            FollowUpClosePulseCompiled = followUpClosePulseCompiled;
            GuardedConfirmFeatureCompiled = guardedConfirmFeatureCompiled;
            GuardedConfirmRuntimeEnabled = guardedConfirmRuntimeEnabled;
            GuardedConfirmEffectiveEnabled = guardedConfirmEffectiveEnabled;
            GuardedConfirmOneShotArmed = guardedConfirmOneShotArmed;
            GuardedConfirmBlockedByPlugin = guardedConfirmBlockedByPlugin;
            PendingType136RetainerMarketOps = pendingType136RetainerMarketOps;
            DisabledForSession = disabledForSession;
            GuardedConfirmDisabledForSession = guardedConfirmDisabledForSession;
            AwaitingAck = awaitingAck;
            GuardedConfirmAwaitingAck = guardedConfirmAwaitingAck;
            InjectedThisListing = injectedThisListing;
            GuardedConfirmClickedThisListing = guardedConfirmClickedThisListing;
            ContextId = contextId;
            GuardedConfirmContextId = guardedConfirmContextId;
            AckTimeoutInMs = ackTimeoutInMs;
            GuardedConfirmAckTimeoutInMs = guardedConfirmAckTimeoutInMs;
        }
    }

    private sealed class PendingListing
    {
        public uint ItemId;
        public int Quantity;
        public int ConfiguredQuantity;
        public uint Price;
    }

    private static readonly Queue<PendingListing> Queue = [];
    private static PendingListing ActiveListing;
    private static bool ActiveContextOpened;
    private static bool ActiveConfirmSent;
    private static int ActiveAttempts;
    private static bool DryRun;
    private static bool AutoConfirm;
    private static bool QueueBuilt;
    private static int SellModeWaitAttempts;
    private static int EnterSellModeAttempts;
    private static bool FailedToEnterSellMode;
    private static bool ManualSellWindowActive;
    private static bool ManualValuesConfigured;
    private static bool WarnedUnsafeSellCallbacks;
    private static bool LoggedManualSellWindowState;
    private static bool LoggedRetainerSellPreConfirmState;
    private static bool PendingSellOpObserved;
    private static uint PendingSellOpContextId;
    private static short PendingSellOpDestSlot;
    private static bool ExperimentalType136Injected;
    private static bool ExperimentalType136FollowUpCloseSent;
    private static long ExperimentalType136InjectedAt;
    private static uint ExperimentalType136ContextId;
    private static bool ExperimentalType136DisabledForSession;
    private static bool ExperimentalType136DisableWarningShown;
    private static bool ExperimentalType136AwaitingAck;
    private static long ExperimentalType136AckTimeoutAt;
    private static int ExperimentalType136PreInjectExactCount;
    private static bool ExperimentalGuardedConfirmClicked;
    private static bool ExperimentalGuardedConfirmAwaitingAck;
    private static bool ExperimentalGuardedConfirmDisabledForSession;
    private static bool ExperimentalGuardedConfirmDisableWarningShown;
    private static long ExperimentalGuardedConfirmAckTimeoutAt;
    private static int ExperimentalGuardedConfirmPreClickExactCount;
    private static uint ExperimentalGuardedConfirmContextId;
    private static bool ExperimentalGuardedConfirmRuntimeEnabled;
    private static bool ExperimentalGuardedConfirmOneShotArmed;
    private static int ExperimentalGuardedConfirmOneShotNoClickAttempts;
    private static long ExperimentalGuardedConfirmOneShotNoClickSinceAt;
    private static long ExperimentalGuardedConfirmSellWindowOpenedAt;
    private static bool ExperimentalGuardedConfirmOneShotSafetyFuseWarningShown;
    private static bool ExperimentalGuardedConfirmBlockedWarningShown;
    private static bool ExperimentalGuardedConfirmPendingOpsWarningShown;
    private static InventoryType ActiveSourceInventoryType;
    private static short ActiveSourceSlot;
    private static int ActiveSourceSlotQuantity;
    private static string PutUpForSaleText;
    private static AdditionalRetainerData PendingConfig;

    // Safety switch: callback-based RetainerSell interactions are currently unstable and can crash the client.
    private const bool UseUnsafeRetainerSellCallbacks = false;

    // Experimental path: directly enqueue observed inventory operation type 136 for RetainerMarket posting.
    private const bool UseExperimentalType136Injection = false;
    private const bool UseExperimentalType136FollowUpClosePulse = false;
    private const int ExperimentalType136AckTimeoutMs = 4000;

    // Experimental path: guarded Confirm button click via AddonMaster wrapper. Kept off by default.
    private const bool UseExperimentalGuardedConfirmClick = false;
    private const bool AllowRuntimeGuardedConfirmOverride = false;
    private const int ExperimentalGuardedConfirmAckTimeoutMs = 4000;
    private const int ExperimentalGuardedConfirmNoClickTimeoutMs = 15000;
    private const int ExperimentalGuardedConfirmWarmupMs = 500;

    private static bool ShouldUseExperimentalGuardedConfirmClick => UseExperimentalGuardedConfirmClick || (AllowRuntimeGuardedConfirmOverride && ExperimentalGuardedConfirmRuntimeEnabled) || ExperimentalGuardedConfirmOneShotArmed;

    // Retainer menu strings differ from inventory context menu text.
    private static readonly string[] MarketSellEntryFallbackTexts =
    [
        "Sell items in your inventory on the market.",
        "Put Up for Sale",
        "Sell items in your retainer's inventory on the market.",
    ];

    private static readonly InventoryType[] SellSourceInventories =
    [
        InventoryType.Inventory1,
        InventoryType.Inventory2,
        InventoryType.Inventory3,
        InventoryType.Inventory4,
        InventoryType.ArmoryMainHand,
        InventoryType.ArmoryHead,
        InventoryType.ArmoryBody,
        InventoryType.ArmoryHands,
        InventoryType.ArmoryLegs,
        InventoryType.ArmoryFeets,
        InventoryType.ArmoryEar,
        InventoryType.ArmoryNeck,
        InventoryType.ArmoryWrist,
        InventoryType.ArmoryRings,
        InventoryType.ArmoryOffHand,
    ];

    private static readonly string[] SellModeAddons =
    [
        "RetainerSellList",
        "RetainerGrid0",
        "RetainerGrid1",
        "RetainerGrid2",
        "RetainerGrid3",
        "RetainerGrid4",
        "RetainerCrystalGrid",
    ];

    internal static void Enqueue(AdditionalRetainerData adata)
    {
        if(!adata.EnableMarketAutoRestock) return;
        if(adata.MarketRestockRules.Count == 0) return;

        ResetState();
        if(P.Memory != null) P.Memory.LogRetainerItemCommandsVerbose = true;
        DryRun = adata.MarketAutoRestockDryRun;
        AutoConfirm = adata.MarketAutoRestockAutoConfirm;
        if(ExperimentalGuardedConfirmOneShotArmed && !AutoConfirm)
        {
            AutoConfirm = true;
            PluginLog.Warning("[MarketRestock] One-shot guarded confirm is armed while AutoConfirm is disabled in config; forcing AutoConfirm=true for this run.");
            DuoLog.Warning("Market restock guarded confirm one-shot is armed: temporarily forcing AutoConfirm for this run.");
        }
        PendingConfig = adata;
        PutUpForSaleText = Svc.Data.GetExcelSheet<Addon>()?.GetRow(99).Text.ToString() ?? "Put Up for Sale";
        PluginLog.Information("[MarketRestock] Enqueued market restock flow for current retainer.");

        P.TaskManager.Enqueue(EnterSellMode, new(timeLimitMS: 1000 * 60, abortOnTimeout: false));
        P.TaskManager.Enqueue(WaitForSellMode, new(timeLimitMS: 1000 * 60, abortOnTimeout: false));
        P.TaskManager.Enqueue(ProcessQueue, new(timeLimitMS: 1000 * 60 * 10, abortOnTimeout: false));
        P.TaskManager.Enqueue(CloseSellMode);
    }

    internal static ExperimentalDebugState GetExperimentalDebugState()
    {
        var now = Environment.TickCount64;
        var remaining = ExperimentalType136AwaitingAck
            ? Math.Max(0, ExperimentalType136AckTimeoutAt - now)
            : 0;
        var confirmRemaining = ExperimentalGuardedConfirmAwaitingAck
            ? Math.Max(0, ExperimentalGuardedConfirmAckTimeoutAt - now)
            : 0;

        return new ExperimentalDebugState(
            UseExperimentalType136Injection,
            UseExperimentalType136FollowUpClosePulse,
            UseExperimentalGuardedConfirmClick,
            ExperimentalGuardedConfirmRuntimeEnabled,
            ShouldUseExperimentalGuardedConfirmClick,
            ExperimentalGuardedConfirmOneShotArmed,
            IsRetainerSellLifecycleKnownConflictPresent(),
            CountPendingType136RetainerMarketOperations(),
            ExperimentalType136DisabledForSession,
            ExperimentalGuardedConfirmDisabledForSession,
            ExperimentalType136AwaitingAck,
            ExperimentalGuardedConfirmAwaitingAck,
            ExperimentalType136Injected,
            ExperimentalGuardedConfirmClicked,
            ExperimentalType136ContextId,
            ExperimentalGuardedConfirmContextId,
            remaining,
            confirmRemaining);
    }

    private static bool IsRetainerSellLifecycleKnownConflictPresent()
    {
        return Svc.PluginInterface.InstalledPlugins.Any(x => x.InternalName == "AllaganMarket" && x.IsLoaded);
    }

    private static int CountPendingType136RetainerMarketOperations()
    {
        try
        {
            var manager = InventoryManager.Instance();
            if(manager == null)
            {
                return 0;
            }

            var count = 0;
            var ops = (InventoryManager.InventoryOperation*)manager;
            for(var i = 0; i < 128; i++)
            {
                var op = ops[i];
                if(op.IsEmpty) continue;
                if(op.Type != 136) continue;
                if(op.DestinationInventoryType != InventoryType.RetainerMarket) continue;
                count++;
            }

            return count;
        }
        catch
        {
            return 0;
        }
    }

    internal static bool SetExperimentalGuardedConfirmRuntimeEnabled(bool enabled)
    {
        if(enabled && !AllowRuntimeGuardedConfirmOverride)
        {
            PluginLog.Warning("[MarketRestock] Experimental guarded confirm runtime toggle was rejected: runtime override is temporarily disabled for stability.");
            DuoLog.Warning("Market restock experimental guarded confirm is temporarily disabled for stability after repeated crashes.");
            ExperimentalGuardedConfirmRuntimeEnabled = false;
            return false;
        }

        if(enabled == ExperimentalGuardedConfirmRuntimeEnabled)
        {
            return false;
        }

        ExperimentalGuardedConfirmRuntimeEnabled = enabled;
        PluginLog.Warning($"[MarketRestock] Experimental guarded confirm runtime toggle set to {enabled}.");
        return true;
    }

    internal static bool ArmExperimentalGuardedConfirmOneShot()
    {
        if(TryGetAddonByName<AtkUnitBase>("RetainerSell", out var sellAddon) && IsAddonReady(sellAddon))
        {
            DuoLog.Warning("Market restock guarded confirm one-shot cannot be armed while RetainerSell is already open.");
            return false;
        }

        if(IsRetainerSellLifecycleKnownConflictPresent())
        {
            DuoLog.Warning("Market restock guarded confirm one-shot arm rejected: AllaganMarket is loaded.");
            return false;
        }

        ExperimentalGuardedConfirmOneShotArmed = true;
        ExperimentalGuardedConfirmOneShotNoClickAttempts = 0;
        ExperimentalGuardedConfirmOneShotNoClickSinceAt = 0;
        PluginLog.Warning("[MarketRestock] Experimental guarded confirm one-shot armed for next eligible listing.");
        DuoLog.Warning("Market restock experimental guarded confirm one-shot armed for next eligible listing.");
        return true;
    }

    internal static void DisarmExperimentalGuardedConfirmOneShot()
    {
        ExperimentalGuardedConfirmOneShotArmed = false;
        ExperimentalGuardedConfirmOneShotNoClickAttempts = 0;
        ExperimentalGuardedConfirmOneShotNoClickSinceAt = 0;
        PluginLog.Warning("[MarketRestock] Experimental guarded confirm one-shot disarmed.");
    }

    internal static void ResetExperimentalSessionSafetyLock()
    {
        ExperimentalType136DisabledForSession = false;
        ExperimentalType136DisableWarningShown = false;
        ExperimentalGuardedConfirmDisabledForSession = false;
        ExperimentalGuardedConfirmDisableWarningShown = false;
        PluginLog.Warning("[MarketRestock] Experimental market restock session safety locks were reset from debug UI.");
    }

    internal static bool IsRetainerSellOpenForDebug()
    {
        return TryGetAddonByName<AtkUnitBase>("RetainerSell", out var sellAddon) && IsAddonReady(sellAddon);
    }

    internal static bool TryDebugRetainerSellConfirmDispatch(int mode)
    {
        if(!TryGetAddonByName<AtkUnitBase>("RetainerSell", out var sellAddon) || !IsAddonReady(sellAddon))
        {
            DuoLog.Warning("Market restock debug confirm probe: RetainerSell is not open.");
            return false;
        }

        if(!EzThrottler.Throttle($"TaskRestockMarketListings.DebugConfirmDispatch.{mode}", 250))
        {
            return false;
        }

        try
        {
            switch(mode)
            {
                case 1:
                    new AddonMaster.RetainerSell(sellAddon).Confirm();
                    PluginLog.Warning("[MarketRestock] Debug confirm probe: AddonMaster.RetainerSell.Confirm().");
                    break;
                case 2:
                    Callback.Fire(sellAddon, true, 0);
                    PluginLog.Warning("[MarketRestock] Debug confirm probe: Callback.Fire(..., 0).");
                    break;
                case 3:
                    Callback.Fire(sellAddon, true, 1);
                    PluginLog.Warning("[MarketRestock] Debug confirm probe: Callback.Fire(..., 1).");
                    break;
                case 4:
                    Callback.Fire(sellAddon, true, 0, (uint)0);
                    PluginLog.Warning("[MarketRestock] Debug confirm probe: Callback.Fire(..., 0, 0u).");
                    break;
                case 5:
                    Callback.Fire(sellAddon, true, 1, (uint)0);
                    PluginLog.Warning("[MarketRestock] Debug confirm probe: Callback.Fire(..., 1, 0u).");
                    break;
                default:
                    DuoLog.Warning($"Market restock debug confirm probe: unknown mode {mode}.");
                    return false;
            }

            DuoLog.Warning($"Market restock debug confirm probe mode {mode} dispatched.");
            return true;
        }
        catch(Exception ex)
        {
            PluginLog.Warning($"[MarketRestock] Debug confirm probe mode {mode} threw: {ex.GetType().Name}: {ex.Message}");
            DuoLog.Warning($"Market restock debug confirm probe mode {mode} threw {ex.GetType().Name}. See log.");
            return false;
        }
    }

    private static bool? EnterSellMode()
    {
        if(IsSellModeReady())
        {
            return true;
        }

        if(EzThrottler.Throttle("TaskRestockMarketListings.EnterSellModeAttemptTick", 250))
        {
            EnterSellModeAttempts++;
            if(EnterSellModeAttempts == 1 || EnterSellModeAttempts % 20 == 0)
            {
                PluginLog.Information($"[MarketRestock] Enter sell mode attempt={EnterSellModeAttempts}");
            }
        }

        if(TryGetAddonMaster<AddonMaster.SelectString>(out var selectString) && selectString.IsAddonReady)
        {
            var best = selectString.Entries
                .Select(x => (Entry: x, Score: ScoreMarketSellEntry(x.Text)))
                .OrderByDescending(x => x.Score)
                .FirstOrDefault();

            if(best.Score > 0)
            {
                if(EzThrottler.Throttle("TaskRestockMarketListings.SelectPutUpForSale", 250))
                {
                    PluginLog.Information($"[MarketRestock] Selecting market sell entry from SelectString: '{best.Entry.Text}'.");
                    best.Entry.Select();
                }
                return false;
            }

            FailedToEnterSellMode = true;
            PluginLog.Warning("[MarketRestock] Put Up for Sale entry not present in SelectString. Skipping this retainer.");
            DuoLog.Warning("Market restock: Put Up for Sale menu entry not available for this retainer now.");
            return true;
        }

        if(TryGetAddonMaster<AddonMaster.SelectIconString>(out var selectIconString) && selectIconString.IsAddonReady)
        {
            var best = selectIconString.Entries
                .Select(x => (Entry: x, Score: ScoreMarketSellEntry(x.Text)))
                .OrderByDescending(x => x.Score)
                .FirstOrDefault();

            if(best.Score > 0)
            {
                if(EzThrottler.Throttle("TaskRestockMarketListings.SelectPutUpForSaleIcon", 250))
                {
                    PluginLog.Information($"[MarketRestock] Selecting market sell entry from SelectIconString: '{best.Entry.Text}'.");
                    best.Entry.Select();
                }
                return false;
            }

            FailedToEnterSellMode = true;
            PluginLog.Warning("[MarketRestock] Put Up for Sale entry not present in SelectIconString. Skipping this retainer.");
            DuoLog.Warning("Market restock: Put Up for Sale menu entry not available for this retainer now.");
            return true;
        }

        if(EnterSellModeAttempts > 80)
        {
            FailedToEnterSellMode = true;
            PluginLog.Warning("[MarketRestock] Retainer menu did not appear in time while entering sell mode. Skipping this retainer.");
            DuoLog.Warning("Market restock: retainer menu did not appear in time.");
            return true;
        }

        Utils.RethrottleGeneric();
        return false;
    }

    private static int ScoreMarketSellEntry(string text)
    {
        if(string.IsNullOrEmpty(text)) return 0;

        // Highest priority: exact player-inventory market option shown in retainer menu.
        if(text.Equals("Sell items in your inventory on the market.")) return 100;

        // Known exact strings in other entry surfaces.
        if(MarketSellEntryFallbackTexts.Contains(text))
        {
            if(text.Equals("Sell items in your retainer's inventory on the market.")) return 20;
            return 80;
        }

        var lower = text.ToLowerInvariant();

        // Prefer options mentioning market + inventory, but strongly avoid retainer-inventory variant.
        var hasMarket = lower.Contains("market");
        var hasInventory = lower.Contains("inventory");
        var isRetainerInventory = lower.Contains("retainer") && hasInventory;

        if(hasMarket && hasInventory && !isRetainerInventory) return 60;
        if(hasMarket && hasInventory) return 10;
        if(hasMarket) return 5;

        return 0;
    }

    private static void BuildQueue(AdditionalRetainerData adata)
    {
        var maxActions = Math.Max(1, adata.MarketAutoRestockMaxListingsPerVisit);
        var enqueued = 0;
        foreach(var rule in adata.MarketRestockRules.Where(x => x.Enabled && x.ItemId > 0 && x.FixedPrice > 0))
        {
            foreach(var target in rule.StackTargets.Where(x => x.Quantity > 0 && x.DesiredListings > 0))
            {
                var existing = CountMarketListings(rule.ItemId, target.Quantity);
                var missing = Math.Max(0, target.DesiredListings - existing);
                for(var i = 0; i < missing; i++)
                {
                    if(enqueued >= maxActions)
                    {
                        DuoLog.Warning($"Market restock: max listings per visit reached ({maxActions}).");
                        return;
                    }

                    if(DryRun)
                    {
                        DuoLog.Information($"> MarketAutoRestock dry run > would list {ExcelItemHelper.GetName(rule.ItemId)} x{target.Quantity} @ {rule.FixedPrice:N0} gil");
                        PluginLog.Information($"[MarketRestock] Dry run would list {ExcelItemHelper.GetName(rule.ItemId)} x{target.Quantity} @ {rule.FixedPrice}.");
                    }
                    else
                    {
                        Queue.Enqueue(new PendingListing
                        {
                            ItemId = rule.ItemId,
                            Quantity = target.Quantity,
                            ConfiguredQuantity = target.Quantity,
                            Price = rule.FixedPrice,
                        });
                        PluginLog.Information($"[MarketRestock] Queue +1 {ExcelItemHelper.GetName(rule.ItemId)} x{target.Quantity} @ {rule.FixedPrice}.");
                    }
                    enqueued++;
                }
            }
        }
        PluginLog.Information($"[MarketRestock] Queue build done. Pending listings={Queue.Count}, dryRun={DryRun}.");
    }

    private static int CountMarketListings(uint itemId, int quantity)
    {
        var count = 0;
        var container = InventoryManager.Instance()->GetInventoryContainer(InventoryType.RetainerMarket);
        if(container == null) return 0;
        for(var i = 0; i < container->Size; i++)
        {
            var slot = container->GetInventorySlot(i);
            if(slot != null && slot->ItemId == itemId && slot->Quantity == quantity)
            {
                count++;
            }
        }
        return count;
    }

    private static bool IsSellModeReady()
    {
        foreach(var name in SellModeAddons)
        {
            if(TryGetAddonByName<AtkUnitBase>(name, out var addon) && IsAddonReady(addon))
            {
                return true;
            }
        }
        return false;
    }

    private static bool? WaitForSellMode()
    {
        if(FailedToEnterSellMode)
        {
            return true;
        }

        if(IsSellModeReady())
        {
            SellModeWaitAttempts = 0;
            return true;
        }

        if(EzThrottler.Throttle("TaskRestockMarketListings.WaitForSellModeAttemptTick", 250))
        {
            SellModeWaitAttempts++;
            if(SellModeWaitAttempts == 1 || SellModeWaitAttempts % 20 == 0)
            {
                PluginLog.Information($"[MarketRestock] Waiting for sell mode... attempt={SellModeWaitAttempts}");
            }
        }

        if(SellModeWaitAttempts > 80)
        {
            FailedToEnterSellMode = true;
            PluginLog.Warning("[MarketRestock] Failed to enter sell mode (RetainerSellList/RetainerGrid not ready). Skipping this retainer.");
            DuoLog.Warning("Market restock: could not enter Put Up for Sale mode. Skipping this retainer.");
            return true;
        }

        Utils.RethrottleGeneric();
        return false;
    }

    private static bool? ProcessQueue()
    {
        if(FailedToEnterSellMode)
        {
            return true;
        }

        if(!IsSellModeReady())
        {
            Utils.RethrottleGeneric();
            return false;
        }

        if(!QueueBuilt)
        {
            BuildQueue(PendingConfig);
            QueueBuilt = true;
            if(Queue.Count == 0)
            {
                PluginLog.Information("[MarketRestock] Queue empty after build. Nothing to do.");
                return true;
            }
        }

        if(ActiveListing == null)
        {
            if(Queue.Count == 0)
            {
                return true;
            }
            ActiveListing = Queue.Dequeue();
            ActiveContextOpened = false;
            ActiveConfirmSent = false;
            ActiveAttempts = 0;
            PluginLog.Information($"[MarketRestock] Processing listing: item={ActiveListing.ItemId}, qty={ActiveListing.ConfiguredQuantity}, price={ActiveListing.Price}. Remaining queue={Queue.Count}.");
        }

        if(TryGetAddonByName<AtkUnitBase>("RetainerSell", out var sellAddon) && IsAddonReady(sellAddon))
        {
            return HandleRetainerSellWindow(sellAddon);
        }

        if(ManualSellWindowActive)
        {
            TrackPendingSellOperation("manual-window-active");
            if(!LoggedManualSellWindowState)
            {
                LogShopEventHandlerState("manual-window-open");
                LogRetainerMarketState("manual-window-open");
                LogPendingInventoryOperations("manual-window-open");
                LoggedManualSellWindowState = true;
            }
            PluginLog.Information("[MarketRestock] Manual RetainerSell window closed; continuing queue.");
            LogShopEventHandlerState("manual-window-close");
            LogRetainerMarketState("manual-window-close");
            LogPendingInventoryOperations("manual-window-close");
            ManualSellWindowActive = false;
            ManualValuesConfigured = false;
            LoggedManualSellWindowState = false;
            LoggedRetainerSellPreConfirmState = false;
            CompleteActiveListing();
            return false;
        }

        if(ActiveConfirmSent)
        {
            if(++ActiveAttempts > 40)
            {
                DuoLog.Warning($"Market restock: confirm was sent for {ExcelItemHelper.GetName(ActiveListing.ItemId)} but sell window did not complete in time.");
                PluginLog.Warning($"[MarketRestock] Confirm sent but did not complete in time: item={ActiveListing.ItemId}.");
            }
            CompleteActiveListing();
            return false;
        }

        if(TryGetAddonMaster<AddonMaster.ContextMenu>(out var contextMenu) && contextMenu.IsAddonReady)
        {
            foreach(var entry in contextMenu.Entries)
            {
                if(entry.Enabled && entry.Text.Equals(PutUpForSaleText))
                {
                    if(EzThrottler.Throttle("TaskRestockMarketListings.SelectPutUpForSale", 250))
                    {
                        entry.Select();
                        ActiveContextOpened = true;
                    }
                    return false;
                }
            }

            DuoLog.Warning($"Market restock: could not find \"{PutUpForSaleText}\" context menu entry for {ExcelItemHelper.GetName(ActiveListing.ItemId)}.");
            PluginLog.Warning($"[MarketRestock] Missing context menu entry '{PutUpForSaleText}' for item={ActiveListing.ItemId}.");
            Callback.Fire(contextMenu.Base, true, -1);
            CompleteActiveListing();
            return false;
        }

        if(!ActiveContextOpened)
        {
            if(!TryGetInventorySlotForPosting(ActiveListing.ItemId, ActiveListing.ConfiguredQuantity, out var sourceType, out var sourceSlot, out var slotQuantity))
            {
                DuoLog.Warning($"Market restock: no stack of {ExcelItemHelper.GetName(ActiveListing.ItemId)} found in player inventory.");
                PluginLog.Warning($"[MarketRestock] No source stack found for item={ActiveListing.ItemId}.");
                CompleteActiveListing();
                return false;
            }

            if(slotQuantity < ActiveListing.ConfiguredQuantity)
            {
                ActiveListing.Quantity = slotQuantity;
                DuoLog.Information($"Market restock: using first available stack for {ExcelItemHelper.GetName(ActiveListing.ItemId)} (configured {ActiveListing.ConfiguredQuantity}, available {slotQuantity}).");
                PluginLog.Information($"[MarketRestock] Fallback to first stack: item={ActiveListing.ItemId}, configured={ActiveListing.ConfiguredQuantity}, available={slotQuantity}.");
            }
            else
            {
                ActiveListing.Quantity = ActiveListing.ConfiguredQuantity;
            }

            if(EzThrottler.Throttle("TaskRestockMarketListings.OpenInventoryContext", 250))
            {
                AgentInventoryContext.Instance()->OpenForItemSlot(sourceType, sourceSlot, 0, 0);
                ActiveContextOpened = true;
                ActiveAttempts++;
                ActiveSourceInventoryType = sourceType;
                ActiveSourceSlot = (short)sourceSlot;
                ActiveSourceSlotQuantity = slotQuantity;
                PluginLog.Information($"[MarketRestock] Opened inventory context: type={sourceType}, slot={sourceSlot}, qty={ActiveListing.Quantity}.");
            }
            return false;
        }

        if(++ActiveAttempts > 20)
        {
            DuoLog.Warning($"Market restock: timed out while trying to list {ExcelItemHelper.GetName(ActiveListing.ItemId)} x{ActiveListing.Quantity}.");
            PluginLog.Warning($"[MarketRestock] Timeout waiting for sell context: item={ActiveListing.ItemId}, qty={ActiveListing.Quantity}.");
            CompleteActiveListing();
        }
        return false;
    }

    private static bool? HandleRetainerSellWindow(AtkUnitBase* sellAddon)
    {
        var master = new AddonMaster.RetainerSell(sellAddon);
        if(!master.IsAddonReady)
        {
            Utils.RethrottleGeneric();
            return false;
        }

        if(!UseUnsafeRetainerSellCallbacks)
        {
            if(ExperimentalGuardedConfirmSellWindowOpenedAt == 0)
            {
                ExperimentalGuardedConfirmSellWindowOpenedAt = Environment.TickCount64;
            }

            ManualSellWindowActive = true;
            var qty = Math.Max(1, ActiveListing.Quantity);
            var price = Math.Max(1, (int)ActiveListing.Price);
            TrackPendingSellOperation("retainer-sell-open");
            if(ApplySafeSellValues(sellAddon, qty, price))
            {
                if(!ManualValuesConfigured)
                {
                    PluginLog.Information($"[MarketRestock] Safely configured RetainerSell values: item={ActiveListing.ItemId}, qty={qty}, price={price}.");
                    if(AutoConfirm)
                    {
                        DuoLog.Information($"Market restock: prefilled {qty} @ {price:N0} gil. Attempting auto-confirm.");
                    }
                    else
                    {
                        DuoLog.Information($"Market restock: prefilled {qty} @ {price:N0} gil. Confirm manually in sell window.");
                    }
                    ManualValuesConfigured = true;
                }
            }

            if(ManualValuesConfigured && !LoggedRetainerSellPreConfirmState)
            {
                LogRetainerMarketState("pre-confirm-window-open");
                LogPendingInventoryOperations("pre-confirm-window-open");
                LoggedRetainerSellPreConfirmState = true;
            }

            if(UseExperimentalType136Injection
                && AutoConfirm
                && ManualValuesConfigured
                && !ExperimentalType136Injected
                && !ExperimentalType136DisabledForSession)
            {
                var injectedContextId = 0u;
                var preInjectExactCount = CountMarketListingsExact(ActiveListing.ItemId, qty, (uint)price);
                if(TryGetPendingType136RetainerMarketOperation(out var staleIndex, out var staleOp))
                {
                    ExperimentalType136DisabledForSession = true;
                    if(!ExperimentalType136DisableWarningShown)
                    {
                        ExperimentalType136DisableWarningShown = true;
                        PluginLog.Warning(
                            $"[MarketRestock] Experimental type136 disabled for this session: stale pending op already exists idx={staleIndex}, ctx={staleOp.ContextId}, src={staleOp.SourceInventoryType}:{staleOp.SourceInventorySlot}, dst={staleOp.DestinationInventoryType}:{staleOp.DestinationInventorySlot}, qty={staleOp.DestinationItemQuantity}, price={staleOp.DestinationItemId}.");
                        DuoLog.Warning("Market restock experimental mode auto-disabled for this session: pending market operation already exists. Restart game/plugin session before retrying experiments.");
                    }
                }
                else
                {
                    ExperimentalType136Injected = TryInjectExperimentalType136Operation(out injectedContextId);
                }

                if(ExperimentalType136Injected)
                {
                    ExperimentalType136ContextId = injectedContextId;
                    ExperimentalType136InjectedAt = Environment.TickCount64;
                    ExperimentalType136AwaitingAck = true;
                    ExperimentalType136AckTimeoutAt = ExperimentalType136InjectedAt + ExperimentalType136AckTimeoutMs;
                    ExperimentalType136PreInjectExactCount = preInjectExactCount;
                    DuoLog.Warning("Market restock experimental mode: submitted type-136 operation. Verify listing result and stability.");
                }
            }

            if(UseExperimentalType136Injection
                && ExperimentalType136Injected
                && ExperimentalType136AwaitingAck)
            {
                if(HandleExperimentalType136AckWait(sellAddon, qty, (uint)price))
                {
                    return false;
                }
            }

            if(AutoConfirm
                && ManualValuesConfigured
                && !ActiveConfirmSent)
            {
                var now = Environment.TickCount64;
                var sellWindowAgeMs = now - ExperimentalGuardedConfirmSellWindowOpenedAt;
                if(sellWindowAgeMs < ExperimentalGuardedConfirmWarmupMs)
                {
                    if(EzThrottler.Throttle("TaskRestockMarketListings.ExperimentalGuardedConfirmWarmup", 1000))
                    {
                        PluginLog.Information($"[MarketRestock] Auto-confirm warmup active: sell window age {sellWindowAgeMs}ms < {ExperimentalGuardedConfirmWarmupMs}ms.");
                    }
                    return false;
                }

                if(EzThrottler.Throttle("TaskRestockMarketListings.AutoConfirmDirectAttempt", 350))
                {
                    try
                    {
                        var dispatchMode = DispatchRetainerSellConfirmPreferred(sellAddon);
                        ActiveConfirmSent = true;
                        ExperimentalGuardedConfirmOneShotArmed = false;
                        ExperimentalGuardedConfirmOneShotNoClickAttempts = 0;
                        ExperimentalGuardedConfirmOneShotNoClickSinceAt = 0;
                        PluginLog.Warning($"[MarketRestock] Auto-confirm attempted via callback mode {dispatchMode}: item={ActiveListing.ItemId}, qty={qty}, price={price}.");
                    }
                    catch(Exception ex)
                    {
                        AutoConfirm = false;
                        PluginLog.Warning($"[MarketRestock] Auto-confirm attempt threw; disabling AutoConfirm for this run: {ex.GetType().Name}: {ex.Message}");
                        DuoLog.Warning("Market restock auto-confirm attempt failed; continuing in manual-confirm mode for this run.");
                    }
                }
            }

            if(ExperimentalGuardedConfirmClicked
                && ExperimentalGuardedConfirmAwaitingAck)
            {
                if(HandleExperimentalGuardedConfirmAckWait(qty, (uint)price))
                {
                    return false;
                }
            }

            if(UseExperimentalType136Injection
                && UseExperimentalType136FollowUpClosePulse
                && ExperimentalType136Injected
                && !ExperimentalType136FollowUpCloseSent)
            {
                var elapsed = Environment.TickCount64 - ExperimentalType136InjectedAt;
                if(elapsed >= 250 && EzThrottler.Throttle("TaskRestockMarketListings.ExperimentalType136ClosePulse", 250))
                {
                    PluginLog.Warning($"[MarketRestock] Experimental type136 follow-up close pulse: ctx={ExperimentalType136ContextId}, elapsedMs={elapsed}");
                    Callback.Fire(sellAddon, true, -1);
                    ExperimentalType136FollowUpCloseSent = true;
                }
            }

            if(!WarnedUnsafeSellCallbacks)
            {
                WarnedUnsafeSellCallbacks = true;
                PluginLog.Warning("[MarketRestock] Unsafe RetainerSell callbacks and confirm click are disabled. Safe numeric setters are active.");
                if(!AutoConfirm)
                {
                    DuoLog.Warning("Market restock safety mode: auto-fills quantity/price safely, but you must confirm manually.");
                }
            }
            return false;
        }

        if(EzThrottler.Throttle("TaskRestockMarketListings.ConfigureListing", 250))
        {
            ActiveAttempts++;
            var qty = Math.Max(1, ActiveListing.Quantity);
            var price = Math.Max(1, (int)ActiveListing.Price);
            if(master.Quantity != qty)
            {
                master.Quantity = qty;
            }
            if(master.AskingPrice != price)
            {
                master.AskingPrice = price;
            }

            if(!AutoConfirm)
            {
                DuoLog.Information($"Market restock: configured listing for {master.ItemName} x{qty} @ {price:N0} gil (awaiting manual confirm).");
                PluginLog.Information($"[MarketRestock] Configured listing awaiting manual confirm: item={ActiveListing.ItemId}, qty={qty}, price={price}.");
                CompleteActiveListing();
                return false;
            }

            if(master.ConfirmButton != null && master.ConfirmButton->IsEnabled)
            {
                master.Confirm();
                ActiveConfirmSent = true;
                ActiveAttempts = 0;
                PluginLog.Information($"[MarketRestock] Confirm clicked: item={ActiveListing.ItemId}, qty={qty}, price={price}.");
            }
            else if(ActiveAttempts > 40)
            {
                DuoLog.Warning($"Market restock: confirm button stayed disabled for {ExcelItemHelper.GetName(ActiveListing.ItemId)} x{qty} @ {price:N0} gil. Skipping this listing.");
                PluginLog.Warning($"[MarketRestock] Confirm button disabled too long: item={ActiveListing.ItemId}, qty={qty}, price={price}.");
                master.Cancel();
                CompleteActiveListing();
            }
        }
        return false;
    }

    private static int DispatchRetainerSellConfirmPreferred(AtkUnitBase* sellAddon)
    {
        try
        {
            Callback.Fire(sellAddon, true, 0);
            return 2;
        }
        catch(Exception ex)
        {
            PluginLog.Warning($"[MarketRestock] Callback confirm mode 2 threw; trying mode 4 fallback: {ex.GetType().Name}: {ex.Message}");
        }

        Callback.Fire(sellAddon, true, 0, (uint)0);
        return 4;
    }

    private static bool HandleExperimentalType136AckWait(AtkUnitBase* sellAddon, int quantity, uint price)
    {
        if(ActiveListing == null)
        {
            ExperimentalType136AwaitingAck = false;
            return false;
        }

        var now = Environment.TickCount64;
        var currentExactCount = CountMarketListingsExact(ActiveListing.ItemId, quantity, price);
        if(currentExactCount > ExperimentalType136PreInjectExactCount)
        {
            PluginLog.Warning(
                $"[MarketRestock] Experimental type136 ack success: ctx={ExperimentalType136ContextId}, exactListingCount {ExperimentalType136PreInjectExactCount} -> {currentExactCount}.");
            ExperimentalType136AwaitingAck = false;
            return false;
        }

        if(now >= ExperimentalType136AckTimeoutAt)
        {
            ExperimentalType136AwaitingAck = false;
            ExperimentalType136DisabledForSession = true;
            if(!ExperimentalType136DisableWarningShown)
            {
                ExperimentalType136DisableWarningShown = true;
                DuoLog.Warning("Market restock experimental mode: no listing ack detected after type-136 injection; experimental path disabled for this session.");
            }

            PluginLog.Warning(
                $"[MarketRestock] Experimental type136 ack timeout: ctx={ExperimentalType136ContextId}, exactListingCount stayed at {currentExactCount}. Closing RetainerSell safely.");
            LogPendingInventoryOperations("experimental-type136-ack-timeout");
            LogRetainerMarketState("experimental-type136-ack-timeout");
            if(EzThrottler.Throttle("TaskRestockMarketListings.ExperimentalType136AckTimeoutClose", 250))
            {
                Callback.Fire(sellAddon, true, -1);
            }
            return true;
        }

        if(EzThrottler.Throttle("TaskRestockMarketListings.ExperimentalType136AckWaitLog", 1000))
        {
            PluginLog.Information(
                $"[MarketRestock] Experimental type136 awaiting ack: ctx={ExperimentalType136ContextId}, exactListingCount={currentExactCount}, deadlineInMs={ExperimentalType136AckTimeoutAt - now}.");
        }

        return true;
    }

    private static bool IsConfirmButtonReady(AddonMaster.RetainerSell master)
    {
        try
        {
            var button = master.ConfirmButton;
            if(button == null)
            {
                return false;
            }

            var node = button->AtkResNode;
            if(node == null)
            {
                return false;
            }

            return button->IsEnabled && node->IsVisible();
        }
        catch(Exception ex)
        {
            if(EzThrottler.Throttle("TaskRestockMarketListings.ConfirmButtonReadException", 2000))
            {
                PluginLog.Warning($"[MarketRestock] Confirm button readiness check failed safely: {ex.GetType().Name}: {ex.Message}");
            }
            return false;
        }
    }

    private static bool HandleExperimentalGuardedConfirmAckWait(int quantity, uint price)
    {
        if(ActiveListing == null)
        {
            ExperimentalGuardedConfirmAwaitingAck = false;
            return false;
        }

        var now = Environment.TickCount64;
        var currentExactCount = CountMarketListingsExact(ActiveListing.ItemId, quantity, price);
        if(currentExactCount > ExperimentalGuardedConfirmPreClickExactCount)
        {
            PluginLog.Warning(
                $"[MarketRestock] Experimental guarded confirm ack success: ctx={ExperimentalGuardedConfirmContextId}, exactListingCount {ExperimentalGuardedConfirmPreClickExactCount} -> {currentExactCount}.");
            ExperimentalGuardedConfirmAwaitingAck = false;
            return false;
        }

        if(now >= ExperimentalGuardedConfirmAckTimeoutAt)
        {
            ExperimentalGuardedConfirmAwaitingAck = false;
            ExperimentalGuardedConfirmDisabledForSession = true;
            if(!ExperimentalGuardedConfirmDisableWarningShown)
            {
                ExperimentalGuardedConfirmDisableWarningShown = true;
                DuoLog.Warning("Market restock experimental guarded confirm: no listing ack detected after click; guarded confirm disabled for this session.");
            }

            PluginLog.Warning(
                $"[MarketRestock] Experimental guarded confirm ack timeout: ctx={ExperimentalGuardedConfirmContextId}, exactListingCount stayed at {currentExactCount}.");
            LogPendingInventoryOperations("experimental-guarded-confirm-ack-timeout");
            LogRetainerMarketState("experimental-guarded-confirm-ack-timeout");
            return false;
        }

        if(EzThrottler.Throttle("TaskRestockMarketListings.ExperimentalGuardedConfirmAckWaitLog", 1000))
        {
            PluginLog.Information(
                $"[MarketRestock] Experimental guarded confirm awaiting ack: ctx={ExperimentalGuardedConfirmContextId}, exactListingCount={currentExactCount}, deadlineInMs={ExperimentalGuardedConfirmAckTimeoutAt - now}.");
        }

        return true;
    }

    private static bool ApplySafeSellValues(AtkUnitBase* sellAddon, int quantity, int price)
    {
        var addon = (AddonRetainerSell*)sellAddon;
        if(addon == null || addon->Quantity == null || addon->AskingPrice == null)
        {
            return false;
        }

        if(EzThrottler.Throttle("TaskRestockMarketListings.SafeSetQuantity", 200))
        {
            addon->Quantity->SetValue(quantity);
        }
        if(EzThrottler.Throttle("TaskRestockMarketListings.SafeSetPrice", 200))
        {
            addon->AskingPrice->SetValue(price);
        }

        // Mark success once values appear on the numeric controls.
        return addon->Quantity->Value == quantity && addon->AskingPrice->Value == price;
    }

    private static void LogShopEventHandlerState(string phase)
    {
        try
        {
            var proxy = ShopEventHandler.AgentProxy.Instance();
            var handler = proxy != null ? proxy->Handler : null;
            if(handler == null)
            {
                PluginLog.Information($"[MarketRestock] ShopEventHandler state ({phase}): handler=null");
                return;
            }

            PluginLog.Information(
                $"[MarketRestock] ShopEventHandler state ({phase}): addonId={proxy->AddonId}, " +
                $"waitingSellConfirm={handler->WaitingForSellConfirm}, waitingTransaction={handler->WaitingForTransactionToFinish}, " +
                $"isTradingWithRetainer={handler->IsTradingWithRetainer}, transactionType={handler->TransactionType}, " +
                $"itemId={handler->TransactionItemId}, count={handler->TransactionItemCount}, sellType={handler->SellInventoryType}, sellSlot={handler->SellInventorySlot}");
        }
        catch(Exception ex)
        {
            PluginLog.Warning($"[MarketRestock] Failed to log ShopEventHandler state for phase={phase}: {ex.Message}");
        }
    }

    private static void LogRetainerMarketState(string phase)
    {
        try
        {
            if(ActiveListing == null)
            {
                PluginLog.Information($"[MarketRestock] RetainerMarket state ({phase}): activeListing=null");
                return;
            }

            var container = InventoryManager.Instance()->GetInventoryContainer(InventoryType.RetainerMarket);
            if(container == null)
            {
                PluginLog.Information($"[MarketRestock] RetainerMarket state ({phase}): container=null");
                return;
            }

            var matches = new List<string>();
            for(var i = 0; i < container->Size; i++)
            {
                var slot = container->GetInventorySlot(i);
                if(slot != null && slot->ItemId == ActiveListing.ItemId)
                {
                    var price = InventoryManager.Instance()->GetRetainerMarketPrice((short)i);
                    matches.Add($"slot={i}, qty={slot->Quantity}, price={price}, flags={slot->Flags}");
                }
            }

            PluginLog.Information($"[MarketRestock] RetainerMarket state ({phase}): item={ActiveListing.ItemId}, configuredQty={ActiveListing.ConfiguredQuantity}, activeQty={ActiveListing.Quantity}, matches={(matches.Count == 0 ? "none" : string.Join(" | ", matches))}");
        }
        catch(Exception ex)
        {
            PluginLog.Warning($"[MarketRestock] Failed to log RetainerMarket state for phase={phase}: {ex.Message}");
        }
    }

    private static void LogPendingInventoryOperations(string phase)
    {
        try
        {
            var manager = InventoryManager.Instance();
            if(manager == null)
            {
                PluginLog.Information($"[MarketRestock] PendingOps ({phase}): manager=null");
                return;
            }

            // InventoryOperation array starts at offset 0 in InventoryManager.
            var ops = (InventoryManager.InventoryOperation*)manager;
            var matches = new List<string>();
            for(var i = 0; i < 128; i++)
            {
                var op = ops[i];
                if(op.IsEmpty) continue;

                matches.Add(
                    $"idx={i}, ctx={op.ContextId}, type={op.Type}, src={op.SourceInventoryType}:{op.SourceInventorySlot} x{op.SourceItemQuantity} item={op.SourceItemId}, dst={op.DestinationInventoryType}:{op.DestinationInventorySlot} x{op.DestinationItemQuantity} item={op.DestinationItemId}");
            }

            PluginLog.Information($"[MarketRestock] PendingOps ({phase}): {(matches.Count == 0 ? "none" : string.Join(" | ", matches))}");
        }
        catch(Exception ex)
        {
            PluginLog.Warning($"[MarketRestock] Failed to log pending inventory operations for phase={phase}: {ex.Message}");
        }
    }

    private static bool TryInjectExperimentalType136Operation(out uint injectedContextId)
    {
        injectedContextId = 0;
        try
        {
            if(ActiveListing == null)
            {
                return false;
            }

            var manager = InventoryManager.Instance();
            if(manager == null)
            {
                PluginLog.Warning("[MarketRestock] Type136 injection aborted: InventoryManager is null.");
                return false;
            }

            if(ActiveSourceInventoryType == InventoryType.Invalid || ActiveSourceSlot < 0)
            {
                PluginLog.Warning("[MarketRestock] Type136 injection aborted: source inventory slot is not initialized.");
                return false;
            }

            var retainerMarket = manager->GetInventoryContainer(InventoryType.RetainerMarket);
            if(retainerMarket == null)
            {
                PluginLog.Warning("[MarketRestock] Type136 injection aborted: RetainerMarket container is null.");
                return false;
            }

            var destinationSlot = -1;
            for(var i = 0; i < retainerMarket->Size; i++)
            {
                var slot = retainerMarket->GetInventorySlot(i);
                if(slot != null && slot->ItemId == 0)
                {
                    destinationSlot = i;
                    break;
                }
            }

            if(destinationSlot < 0)
            {
                PluginLog.Warning("[MarketRestock] Type136 injection aborted: no empty RetainerMarket slot found.");
                return false;
            }

            var ops = (InventoryManager.InventoryOperation*)manager;
            var pendingIndex = -1;
            for(var i = 0; i < 128; i++)
            {
                if(ops[i].IsEmpty)
                {
                    pendingIndex = i;
                    break;
                }
            }

            if(pendingIndex < 0)
            {
                PluginLog.Warning("[MarketRestock] Type136 injection aborted: no empty pending operation slot.");
                return false;
            }

            var contextId = manager->NextContextId;
            manager->NextContextId++;

            var op = new InventoryManager.InventoryOperation
            {
                IsEmpty = false,
                ContextId = contextId,
                Type = 136,
                SourceInventoryType = ActiveSourceInventoryType,
                SourceInventorySlot = ActiveSourceSlot,
                SourceItemQuantity = ActiveSourceSlotQuantity,
                SourceItemId = 0,
                DestinationInventoryType = InventoryType.RetainerMarket,
                DestinationInventorySlot = (short)destinationSlot,
                DestinationItemQuantity = ActiveListing.Quantity,
                DestinationItemId = ActiveListing.Price,
            };

            ops[pendingIndex] = op;
            injectedContextId = contextId;

            PluginLog.Warning(
                $"[MarketRestock] Experimental type136 injected: pendingIdx={pendingIndex}, ctx={contextId}, src={op.SourceInventoryType}:{op.SourceInventorySlot} x{op.SourceItemQuantity}, dst={op.DestinationInventoryType}:{op.DestinationInventorySlot} x{op.DestinationItemQuantity}, price={op.DestinationItemId}");
            return true;
        }
        catch(Exception ex)
        {
            PluginLog.Warning($"[MarketRestock] Type136 injection exception: {ex.Message}");
            return false;
        }
    }

    private static bool TryGetPendingType136RetainerMarketOperation(out int index, out InventoryManager.InventoryOperation operation)
    {
        index = -1;
        operation = default;
        try
        {
            var manager = InventoryManager.Instance();
            if(manager == null)
            {
                return false;
            }

            var ops = (InventoryManager.InventoryOperation*)manager;
            for(var i = 0; i < 128; i++)
            {
                var op = ops[i];
                if(op.IsEmpty) continue;
                if(op.Type != 136) continue;
                if(op.DestinationInventoryType != InventoryType.RetainerMarket) continue;

                index = i;
                operation = op;
                return true;
            }

            return false;
        }
        catch(Exception ex)
        {
            PluginLog.Warning($"[MarketRestock] Failed to inspect pending type136 operations: {ex.Message}");
            return false;
        }
    }

    private static int CountMarketListingsExact(uint itemId, int quantity, uint price)
    {
        var count = 0;
        var container = InventoryManager.Instance()->GetInventoryContainer(InventoryType.RetainerMarket);
        if(container == null)
        {
            return 0;
        }

        for(var i = 0; i < container->Size; i++)
        {
            var slot = container->GetInventorySlot(i);
            if(slot == null || slot->ItemId != itemId || slot->Quantity != quantity)
            {
                continue;
            }

            if(InventoryManager.Instance()->GetRetainerMarketPrice((short)i) == price)
            {
                count++;
            }
        }
        return count;
    }

    private static void TrackPendingSellOperation(string phase)
    {
        if(ActiveListing == null) return;

        try
        {
            var manager = InventoryManager.Instance();
            if(manager == null)
            {
                return;
            }

            var ops = (InventoryManager.InventoryOperation*)manager;
            var found = false;
            var foundIndex = -1;
            var foundOp = default(InventoryManager.InventoryOperation);
            var foundConfidence = "low";

            for(var i = 0; i < 128; i++)
            {
                var op = ops[i];
                if(op.IsEmpty) continue;
                if(op.Type != 136) continue;
                if(op.DestinationInventoryType != InventoryType.RetainerMarket) continue;

                // Confidence tags help us distinguish likely current listing ops from unrelated market operations.
                foundConfidence = "low";
                if(op.DestinationItemQuantity == ActiveListing.Quantity && (uint)op.DestinationItemId == ActiveListing.Price)
                {
                    foundConfidence = "high";
                }
                else if(op.SourceInventoryType != InventoryType.Invalid && op.SourceInventorySlot >= 0)
                {
                    foundConfidence = "medium";
                }

                found = true;
                foundIndex = i;
                foundOp = op;
                break;
            }

            if(found)
            {
                if(!PendingSellOpObserved)
                {
                    PendingSellOpObserved = true;
                    PendingSellOpContextId = foundOp.ContextId;
                    PendingSellOpDestSlot = foundOp.DestinationInventorySlot;
                    PluginLog.Information(
                        $"[MarketRestock] PendingSellOp queued ({phase}): confidence={foundConfidence}, idx={foundIndex}, ctx={foundOp.ContextId}, type={foundOp.Type}, src={foundOp.SourceInventoryType}:{foundOp.SourceInventorySlot} x{foundOp.SourceItemQuantity}, dst={foundOp.DestinationInventoryType}:{foundOp.DestinationInventorySlot} x{foundOp.DestinationItemQuantity}, price={foundOp.DestinationItemId}");
                }
                else if(PendingSellOpContextId != foundOp.ContextId || PendingSellOpDestSlot != foundOp.DestinationInventorySlot)
                {
                    PluginLog.Information(
                        $"[MarketRestock] PendingSellOp changed ({phase}): confidence={foundConfidence}, oldCtx={PendingSellOpContextId}, newCtx={foundOp.ContextId}, oldDstSlot={PendingSellOpDestSlot}, newDstSlot={foundOp.DestinationInventorySlot}");
                    PendingSellOpContextId = foundOp.ContextId;
                    PendingSellOpDestSlot = foundOp.DestinationInventorySlot;
                }
            }
            else if(PendingSellOpObserved)
            {
                PluginLog.Information($"[MarketRestock] PendingSellOp cleared ({phase}): lastCtx={PendingSellOpContextId}, lastDstSlot={PendingSellOpDestSlot}");
                PendingSellOpObserved = false;
                PendingSellOpContextId = 0;
                PendingSellOpDestSlot = -1;
            }
            else
            {
                if(EzThrottler.Throttle($"TaskRestockMarketListings.PendingSellOpNone.{phase}", 1000))
                {
                    PluginLog.Information($"[MarketRestock] PendingSellOp none ({phase}): no type=136 RetainerMarket operation currently visible.");
                }
            }
        }
        catch(Exception ex)
        {
            PluginLog.Warning($"[MarketRestock] Failed to track pending sell operation for phase={phase}: {ex.Message}");
        }
    }

    private static bool TryGetInventorySlotForPosting(uint itemId, int requiredQuantity, out InventoryType inventoryType, out int slot, out int slotQuantity)
    {
        var minQuantity = Math.Max(1, requiredQuantity);
        var hasFallback = false;
        var fallbackType = InventoryType.Invalid;
        var fallbackSlot = -1;
        var fallbackQuantity = 0;

        foreach(var type in SellSourceInventories)
        {
            var inv = InventoryManager.Instance()->GetInventoryContainer(type);
            if(inv == null) continue;
            for(var i = 0; i < inv->Size; i++)
            {
                var item = inv->GetInventorySlot(i);
                if(item != null && item->ItemId == itemId)
                {
                    if(!hasFallback)
                    {
                        hasFallback = true;
                        fallbackType = type;
                        fallbackSlot = i;
                        fallbackQuantity = item->Quantity;
                    }

                    if(item->Quantity >= minQuantity)
                    {
                        inventoryType = type;
                        slot = i;
                        slotQuantity = item->Quantity;
                        return true;
                    }
                }
            }
        }

        if(hasFallback)
        {
            inventoryType = fallbackType;
            slot = fallbackSlot;
            slotQuantity = fallbackQuantity;
            return true;
        }

        inventoryType = InventoryType.Invalid;
        slot = -1;
        slotQuantity = 0;
        return false;
    }

    private static bool? CloseSellMode()
    {
        if(TryGetAddonByName<AtkUnitBase>("RetainerSell", out var sellAddon) && IsAddonReady(sellAddon))
        {
            if(EzThrottler.Throttle("TaskRestockMarketListings.CloseRetainerSell", 250))
            {
                Callback.Fire(sellAddon, true, -1);
            }
            return false;
        }

        foreach(var addonName in SellModeAddons)
        {
            if(TryGetAddonByName<AtkUnitBase>(addonName, out var addon) && IsAddonReady(addon))
            {
                if(EzThrottler.Throttle($"TaskRestockMarketListings.Close.{addonName}", 250))
                {
                    Callback.Fire(addon, true, -1);
                }
                return false;
            }
        }

        if(P.Memory != null) P.Memory.LogRetainerItemCommandsVerbose = false;
        return true;
    }

    private static void CompleteActiveListing()
    {
        ActiveListing = null;
        ActiveContextOpened = false;
        ActiveConfirmSent = false;
        ActiveAttempts = 0;
        PendingSellOpObserved = false;
        PendingSellOpContextId = 0;
        PendingSellOpDestSlot = -1;
        ExperimentalType136Injected = false;
        ExperimentalType136FollowUpCloseSent = false;
        ExperimentalType136InjectedAt = 0;
        ExperimentalType136ContextId = 0;
        ExperimentalType136AwaitingAck = false;
        ExperimentalType136AckTimeoutAt = 0;
        ExperimentalType136PreInjectExactCount = 0;
        ExperimentalGuardedConfirmClicked = false;
        ExperimentalGuardedConfirmAwaitingAck = false;
        ExperimentalGuardedConfirmAckTimeoutAt = 0;
        ExperimentalGuardedConfirmPreClickExactCount = 0;
        ExperimentalGuardedConfirmContextId = 0;
        ExperimentalGuardedConfirmOneShotNoClickAttempts = 0;
        ExperimentalGuardedConfirmOneShotNoClickSinceAt = 0;
        ExperimentalGuardedConfirmSellWindowOpenedAt = 0;
        ExperimentalGuardedConfirmOneShotSafetyFuseWarningShown = false;
        ExperimentalGuardedConfirmBlockedWarningShown = false;
        ExperimentalGuardedConfirmPendingOpsWarningShown = false;
        ActiveSourceInventoryType = InventoryType.Invalid;
        ActiveSourceSlot = -1;
        ActiveSourceSlotQuantity = 0;
    }

    private static void ResetState()
    {
        Queue.Clear();
        CompleteActiveListing();
        if(P.Memory != null) P.Memory.LogRetainerItemCommandsVerbose = false;
        DryRun = false;
        AutoConfirm = true;
        QueueBuilt = false;
        SellModeWaitAttempts = 0;
        EnterSellModeAttempts = 0;
        FailedToEnterSellMode = false;
        ManualSellWindowActive = false;
        ManualValuesConfigured = false;
        WarnedUnsafeSellCallbacks = false;
        LoggedManualSellWindowState = false;
        LoggedRetainerSellPreConfirmState = false;
        PendingConfig = null;
        PutUpForSaleText = "Put Up for Sale";
    }
}
