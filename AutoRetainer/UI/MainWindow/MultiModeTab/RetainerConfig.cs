using AutoRetainerAPI;
using AutoRetainerAPI.Configuration;
using ECommons.ExcelServices;
using System;

namespace AutoRetainer.UI.MainWindow.MultiModeTab;
public static unsafe class RetainerConfig
{
    public static void Draw(OfflineRetainerData ret, OfflineCharacterData data, AdditionalRetainerData adata)
    {
        ImGui.CollapsingHeader($"{Censor.Retainer(ret.Name)} - {Censor.Character(data.Name)} Configuration  ##conf", ImGuiTreeNodeFlags.DefaultOpen | ImGuiTreeNodeFlags.Bullet | ImGuiTreeNodeFlags.OpenOnArrow);
        ImGuiEx.Text($"Additional Post-venture Tasks:");
        //ImGui.Checkbox($"Entrust Duplicates", ref adata.EntrustDuplicates);
        var selectedPlan = C.EntrustPlans.FirstOrDefault(x => x.Guid == adata.EntrustPlan);
        ImGuiEx.TextV($"Entrust Items:");
        if(!C.EnableEntrustManager) ImGuiEx.HelpMarker("Globally disabled in settings", EColor.RedBright, FontAwesomeIcon.ExclamationTriangle.ToIconString());
        ImGui.SameLine();
        ImGui.SetNextItemWidth(150f);
        if(ImGui.BeginCombo($"##select", selectedPlan?.Name ?? "Disabled", ImGuiComboFlags.HeightLarge))
        {
            if(ImGui.Selectable("Disabled")) adata.EntrustPlan = Guid.Empty;
            for(var i = 0; i < C.EntrustPlans.Count; i++)
            {
                var plan = C.EntrustPlans[i];
                ImGui.PushID(plan.Guid.ToString());
                if(ImGui.Selectable(plan.Name, plan == selectedPlan))
                {
                    adata.EntrustPlan = plan.Guid;
                }
                ImGui.PopID();
            }
            ImGui.EndCombo();
        }
        if(ImGuiEx.IconButtonWithText(FontAwesomeIcon.Copy, "Copy entrust plan to..."))
        {
            ImGui.OpenPopup($"CopyEntrustPlanTo");
        }
        if(ImGui.BeginPopup("CopyEntrustPlanTo"))
        {
            if(ImGui.Selectable("To all other retainers of this character"))
            {
                var cnt = 0;
                foreach(var x in data.RetainerData)
                {
                    cnt++;
                    Utils.GetAdditionalData(data.CID, x.Name).EntrustPlan = adata.EntrustPlan;
                }
                Notify.Info($"Changed {cnt} retainers");
            }
            if(ImGui.Selectable("To all other retainers without entrust plan of this character"))
            {
                foreach(var x in data.RetainerData)
                {
                    var cnt = 0;
                    if(!C.EntrustPlans.Any(s => s.Guid == adata.EntrustPlan))
                    {
                        Utils.GetAdditionalData(data.CID, x.Name).EntrustPlan = adata.EntrustPlan;
                        cnt++;
                    }
                    Notify.Info($"Changed {cnt} retainers");
                }
            }
            if(ImGui.Selectable("To all other retainers of ALL characters"))
            {
                var cnt = 0;
                foreach(var offlineData in C.OfflineData)
                {
                    foreach(var x in offlineData.RetainerData)
                    {
                        Utils.GetAdditionalData(offlineData.CID, x.Name).EntrustPlan = adata.EntrustPlan;
                        cnt++;
                    }
                }
                Notify.Info($"Changed {cnt} retainers");
            }
            if(ImGui.Selectable("To all other retainers without entrust plan of ALL characters"))
            {
                var cnt = 0;
                foreach(var offlineData in C.OfflineData)
                {
                    foreach(var x in offlineData.RetainerData)
                    {
                        var a = Utils.GetAdditionalData(data.CID, x.Name);
                        if(!C.EntrustPlans.Any(s => s.Guid == a.EntrustPlan))
                        {
                            a.EntrustPlan = adata.EntrustPlan;
                            cnt++;
                        }
                    }
                }
                Notify.Info($"Changed {cnt} retainers");
            }
            ImGui.EndPopup();
        }
        ImGui.Checkbox($"Withdraw/Deposit Gil", ref adata.WithdrawGil);
        if(adata.WithdrawGil)
        {
            if(ImGui.RadioButton("Withdraw", !adata.Deposit)) adata.Deposit = false;
            if(ImGui.RadioButton("Deposit", adata.Deposit)) adata.Deposit = true;
            ImGuiEx.SetNextItemWidthScaled(200f);
            ImGui.InputInt($"Amount, %", ref adata.WithdrawGilPercent.ValidateRange(1, 100), 1, 10);
        }

        ImGui.Separator();
        ImGuiEx.TextV("Market Auto Restock:");
        ImGui.Checkbox("Enable market auto restock", ref adata.EnableMarketAutoRestock);
        if(adata.EnableMarketAutoRestock)
        {
            ImGui.Checkbox("Dry run", ref adata.MarketAutoRestockDryRun);
            ImGui.SameLine();
            ImGui.Checkbox("Auto confirm listing", ref adata.MarketAutoRestockAutoConfirm);
            ImGuiEx.SetNextItemWidthScaled(200f);
            ImGui.InputInt("Max listings per visit", ref adata.MarketAutoRestockMaxListingsPerVisit.ValidateRange(1, 100), 1, 5);

            for(var i = 0; i < adata.MarketRestockRules.Count; i++)
            {
                var rule = adata.MarketRestockRules[i];
                ImGui.PushID($"MarketRule{i}");
                ImGui.Separator();
                ImGui.Checkbox("Enabled", ref rule.Enabled);
                ImGui.SameLine();
                if(ImGuiEx.IconButton(FontAwesomeIcon.Trash))
                {
                    adata.MarketRestockRules.RemoveAt(i);
                    ImGui.PopID();
                    break;
                }

                var itemId = (int)rule.ItemId;
                ImGui.SetNextItemWidth(180f);
                ImGui.InputInt("Item ID", ref itemId, 0, 0);
                rule.ItemId = (uint)Math.Max(0, itemId);

                var fixedPrice = (int)rule.FixedPrice;
                ImGui.SetNextItemWidth(180f);
                ImGui.InputInt("Fixed Price", ref fixedPrice.ValidateRange(1, int.MaxValue), 100, 1000);
                rule.FixedPrice = (uint)Math.Max(1, fixedPrice);

                if(rule.ItemId > 0)
                {
                    ImGuiEx.Text($"Item: {ExcelItemHelper.GetName(rule.ItemId)}");
                }

                ImGuiEx.Text("Stack Targets");
                for(var s = 0; s < rule.StackTargets.Count; s++)
                {
                    var stack = rule.StackTargets[s];
                    ImGui.PushID($"Stack{s}");
                    ImGui.SetNextItemWidth(120f);
                    ImGui.InputInt("Qty", ref stack.Quantity.ValidateRange(1, 999), 1, 5);
                    ImGui.SameLine();
                    ImGui.SetNextItemWidth(120f);
                    ImGui.InputInt("Listings", ref stack.DesiredListings.ValidateRange(1, 20), 1, 2);
                    ImGui.SameLine();
                    if(ImGuiEx.IconButton(FontAwesomeIcon.Trash))
                    {
                        rule.StackTargets.RemoveAt(s);
                        ImGui.PopID();
                        break;
                    }
                    ImGui.PopID();
                }

                if(ImGuiEx.IconButtonWithText(FontAwesomeIcon.Plus, "Add Stack Target"))
                {
                    rule.StackTargets.Add(new()
                    {
                        Quantity = 1,
                        DesiredListings = 1,
                    });
                }

                ImGui.PopID();
            }

            if(ImGuiEx.IconButtonWithText(FontAwesomeIcon.Plus, "Add Restock Rule"))
            {
                adata.MarketRestockRules.Add(new()
                {
                    Enabled = true,
                    ItemId = 0,
                    FixedPrice = 1,
                    StackTargets = [new() { Quantity = 1, DesiredListings = 1 }],
                });
            }
        }
        ImGui.Separator();
        Svc.PluginInterface.GetIpcProvider<ulong, string, object>(ApiConsts.OnRetainerSettingsDraw).SendMessage(data.CID, ret.Name);
        if(C.Verbose)
        {
            if(ImGui.Button("Fake ready"))
            {
                ret.VentureEndsAt = 1;
            }
            if(ImGui.Button("Fake unready"))
            {
                ret.VentureEndsAt = P.Time + 60 * 60;
            }
        }
    }
}
