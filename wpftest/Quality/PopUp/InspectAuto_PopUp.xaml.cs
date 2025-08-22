using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;

namespace WizMes_WellMade.Quality.PopUp
{
    public partial class InspectAuto_PopUp : Window
    {
        Lib lib = new Lib();

        // 파일 경로
        private string settingsFilePath = "CellSettings.json";
        private string presetFilePath = "CellSettingsWithPresets.json";
        private List<CellSettingPreset> presets = new List<CellSettingPreset>();

        // 현재 선택된 프리셋 저장용
        private CellSettingPreset currentSelectedPreset = null;

        public InspectAuto_PopUp()
        {
            InitializeComponent();
            RegisterTextBoxEvents();
            LoadPresetsAndSettings();
        }

        #region 설정 관리

        // 프리셋과 설정을 모두 로드
        private void LoadPresetsAndSettings()
        {
            try
            {
                // 프리셋 파일이 있으면 우선 로드
                if (File.Exists(presetFilePath))
                {
                    string json = File.ReadAllText(presetFilePath);
                    var collection = JsonConvert.DeserializeObject<PresetCollection>(json);

                    if (collection != null)
                    {
                        // 현재 설정 로드
                        if (collection.CurrentSettings != null)
                        {
                            ApplySettingsToUI(collection.CurrentSettings);
                        }

                        // 프리셋 목록 로드
                        presets = collection.Presets ?? new List<CellSettingPreset>();

                        // 마지막 선택된 프리셋 복원
                        if (!string.IsNullOrEmpty(collection.LastSelectedPresetName))
                        {
                            currentSelectedPreset = presets.FirstOrDefault(p => p.Name == collection.LastSelectedPresetName);
                        }
                    }
                }
                // 기존 파일만 있으면 기존 방식으로 로드
                else if (File.Exists(settingsFilePath))
                {
                    LoadSettings();
                }

                // 프리셋이 없으면 기본 프리셋 생성
                if (presets.Count == 0)
                {
                    CreateDefaultPreset();
                }

                RefreshPresetComboBox();

                // 마지막 선택된 프리셋이 있으면 선택, 없으면 첫 번째 프리셋 선택
                if (currentSelectedPreset != null && presets.Contains(currentSelectedPreset))
                {
                    cboPresets.SelectedItem = currentSelectedPreset;
                }
                else if (presets.Count > 0)
                {
                    cboPresets.SelectedIndex = 0;
                    currentSelectedPreset = presets[0];
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"설정 로드 중 오류가 발생했습니다: {ex.Message}", "오류");
            }
        }

        // 기본 프리셋 생성
        private void CreateDefaultPreset()
        {
            var defaultPreset = new CellSettingPreset
            {
                Name = "새 프리셋",
                Settings = new CellSettings(), // 모든 값이 기본값(false, "")인 상태
                CreatedDate = DateTime.Now
            };

            presets.Add(defaultPreset);
            SavePresetsAndSettings(); // 기본 프리셋을 파일에 저장
        }

        // 설정을 UI에 적용하는 공통 메서드
        private void ApplySettingsToUI(CellSettings settings)
        {
            LoadSettingItem(chkLotNo, txtLotNo, settings.LotNo);
            LoadSettingItem(chkModelID, txtModelID, settings.ModelID);
            LoadSettingItem(chkBuyerArticleNo, txtBuyerArticleNo, settings.BuyerArticleNo);
            LoadSettingItem(chkArticleID, txtArticleID, settings.ArticleID);
            LoadSettingItem(chkInspectDate, txtInspectDate, settings.InspectDate);
            LoadSettingItem(chkName, txtName, settings.Name);
            LoadSettingItem(chkProcessID, txtProcessID, settings.ProcessID);
            LoadSettingItem(chkMachineID, txtMachineID, settings.MachineID);
            LoadSettingItem(chkInspectLevel, txtInspectLevel, settings.InspectLevel);
            LoadSettingItem(chkIRELevel, txtIRELevel, settings.IRELevel);
            LoadSettingItem(chkCustomID, txtCustomID, settings.CustomID);
            LoadSettingItem(chkInOutDate, txtInOutDate, settings.InOutDate);
            LoadSettingItem(chkFMLGubun, txtFMLGubun, settings.FMLGubun);
            LoadSettingItem(chkSumInspectQty, txtSumInspectQty, settings.SumInspectQty);
            LoadSettingItem(chkDefectYN, txtDefectYN, settings.DefectYN);
            LoadSettingItem(chkSumDefectQty, txtSumDefectQty, settings.SumDefectQty);
        }

        // 현재 UI에서 설정값들을 가져오기
        private CellSettings GetCurrentSettings()
        {
            return new CellSettings
            {
                LotNo = new CellSettingItem { Checked = chkLotNo.IsChecked == true, Value = txtLotNo.Text },
                ModelID = new CellSettingItem { Checked = chkModelID.IsChecked == true, Value = txtModelID.Text },
                BuyerArticleNo = new CellSettingItem { Checked = chkBuyerArticleNo.IsChecked == true, Value = txtBuyerArticleNo.Text },
                ArticleID = new CellSettingItem { Checked = chkArticleID.IsChecked == true, Value = txtArticleID.Text },
                InspectDate = new CellSettingItem { Checked = chkInspectDate.IsChecked == true, Value = txtInspectDate.Text },
                Name = new CellSettingItem { Checked = chkName.IsChecked == true, Value = txtName.Text },
                ProcessID = new CellSettingItem { Checked = chkProcessID.IsChecked == true, Value = txtProcessID.Text },
                MachineID = new CellSettingItem { Checked = chkMachineID.IsChecked == true, Value = txtMachineID.Text },
                InspectLevel = new CellSettingItem { Checked = chkInspectLevel.IsChecked == true, Value = txtInspectLevel.Text },
                IRELevel = new CellSettingItem { Checked = chkIRELevel.IsChecked == true, Value = txtIRELevel.Text },
                CustomID = new CellSettingItem { Checked = chkCustomID.IsChecked == true, Value = txtCustomID.Text },
                InOutDate = new CellSettingItem { Checked = chkInOutDate.IsChecked == true, Value = txtInOutDate.Text },
                FMLGubun = new CellSettingItem { Checked = chkFMLGubun.IsChecked == true, Value = txtFMLGubun.Text },
                SumInspectQty = new CellSettingItem { Checked = chkSumInspectQty.IsChecked == true, Value = txtSumInspectQty.Text },
                DefectYN = new CellSettingItem { Checked = chkDefectYN.IsChecked == true, Value = txtDefectYN.Text },
                SumDefectQty = new CellSettingItem { Checked = chkSumDefectQty.IsChecked == true, Value = txtSumDefectQty.Text }
            };
        }

        // 기존 설정 저장 (호환성 유지)
        private void SaveSettings()
        {
            try
            {
                var settings = GetCurrentSettings();
                string json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                File.WriteAllText(settingsFilePath, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"설정 저장 중 오류가 발생했습니다: {ex.Message}", "오류");
            }
        }

        // 기존 설정 로드 (호환성 유지)
        private void LoadSettings()
        {
            try
            {
                if (File.Exists(settingsFilePath))
                {
                    string json = File.ReadAllText(settingsFilePath);
                    var settings = JsonConvert.DeserializeObject<CellSettings>(json);

                    if (settings != null)
                    {
                        ApplySettingsToUI(settings);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"설정 로드 중 오류가 발생했습니다: {ex.Message}", "오류");
            }
        }

        #endregion

        #region 프리셋 관리

        // 프리셋 콤보박스 새로고침
        private void RefreshPresetComboBox()
        {
            cboPresets.Items.Clear();
            foreach (var preset in presets)
            {
                cboPresets.Items.Add(preset);
            }
        }

        // 프리셋과 현재 설정을 파일에 저장
        private void SavePresetsAndSettings()
        {
            try
            {
                var collection = new PresetCollection
                {
                    CurrentSettings = GetCurrentSettings(),
                    Presets = presets,
                    LastSelectedPresetName = currentSelectedPreset?.Name ?? "" // 현재 선택된 프리셋 이름 저장
                };

                string json = JsonConvert.SerializeObject(collection, Formatting.Indented);
                File.WriteAllText(presetFilePath, json);

                // 기존 파일도 저장 (호환성)
                SaveSettings();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"프리셋 저장 중 오류가 발생했습니다: {ex.Message}", "오류");
            }
        }

        #endregion

        #region 프리셋 이벤트 핸들러

        // 프리셋 선택 시
        private void cboPresets_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // 프로그래밍 방식으로 선택이 변경되는 경우 무시 (무한 루프 방지)
            if (e.AddedItems.Count == 0)
                return;

            if (cboPresets.SelectedItem is CellSettingPreset selectedPreset)
            {
                try
                {
                    // 현재 선택된 프리셋 업데이트
                    currentSelectedPreset = selectedPreset;

                    // 선택된 프리셋의 설정을 UI에 적용
                    ApplySettingsToUI(selectedPreset.Settings);

                    // ComboBox 텍스트를 프리셋 이름으로 설정
                    cboPresets.Text = selectedPreset.Name;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"프리셋 적용 중 오류가 발생했습니다: {ex.Message}", "오류");
                }
            }
        }

        // 새 프리셋 저장 버튼 (현재 설정으로 새 프리셋 생성)
        private void btnSaveNewPreset_Click(object sender, RoutedEventArgs e)
        {
            string presetName = cboPresets.Text?.Trim();

            if (string.IsNullOrEmpty(presetName))
            {
                MessageBox.Show("프리셋 이름을 입력해주세요.");
                cboPresets.Focus();
                return;
            }

            // 이름 중복 확인
            var existing = presets.FirstOrDefault(p => p.Name == presetName);
            if (existing != null)
            {
                if (MessageBox.Show($"'{presetName}' 프리셋이 이미 존재합니다. 덮어쓰시겠습니까?",
                                   "확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    // 기존 프리셋 덮어쓰기
                    existing.Settings = GetCurrentSettings();
                    existing.CreatedDate = DateTime.Now;

                    // 덮어쓴 프리셋 선택
                    cboPresets.SelectedItem = existing;
                }
                else
                {
                    return;
                }
            }
            else
            {
                // 새 프리셋 생성
                var newPreset = new CellSettingPreset
                {
                    Name = presetName,
                    Settings = GetCurrentSettings()
                };
                presets.Add(newPreset);

                RefreshPresetComboBox();

                // 새로 만든 프리셋 선택
                cboPresets.SelectedItem = newPreset;
            }

            SavePresetsAndSettings();
            MessageBox.Show("새 프리셋이 저장되었습니다.");
        }

        // 선택 항목 삭제 버튼
        private void btnDeletePreset_Click(object sender, RoutedEventArgs e)
        {
            if (cboPresets.SelectedItem is CellSettingPreset preset)
            {
                if (MessageBox.Show($"'{preset.Name}' 프리셋을 삭제하시겠습니까?",
                                   "확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    // 현재 선택된 프리셋 인덱스 저장
                    int currentIndex = cboPresets.SelectedIndex;

                    // 프리셋 삭제
                    presets.Remove(preset);

                    // 삭제된 프리셋이 현재 선택된 프리셋이었다면 null로 설정
                    if (currentSelectedPreset == preset)
                    {
                        currentSelectedPreset = null;
                    }

                    RefreshPresetComboBox();

                    // 자동 선택 로직
                    if (presets.Count > 0)
                    {
                        CellSettingPreset presetToSelect = null;

                        // 가장 최신 프리셋 찾기 (CreatedDate 기준)
                        presetToSelect = presets.OrderByDescending(p => p.CreatedDate).First();

                        // 최신 프리셋 선택
                        cboPresets.SelectedItem = presetToSelect;
                        currentSelectedPreset = presetToSelect;

                        // 선택된 프리셋의 설정 적용
                        ApplySettingsToUI(presetToSelect.Settings);
                    }
                    else
                    {
                        // 프리셋이 모두 삭제되면 기본 프리셋 생성
                        CreateDefaultPreset();
                        RefreshPresetComboBox();
                        if (presets.Count > 0)
                        {
                            cboPresets.SelectedIndex = 0;
                            currentSelectedPreset = presets[0];
                        }
                    }

                    SavePresetsAndSettings();
                    MessageBox.Show("프리셋이 삭제되었습니다.\n가장 최신 프리셋이 선택되었습니다.");
                }
            }
            else
            {
                MessageBox.Show("삭제할 프리셋을 선택해주세요.");
            }
        }

        // 우클릭 메뉴 - 삭제 (기존 코드 호환)
        private void MenuItem_DeletePreset_Click(object sender, RoutedEventArgs e)
        {
            btnDeletePreset_Click(sender, e);
        }

        #endregion

        #region 기존 이벤트 핸들러들

        // 개별 설정 항목 로드 (체크박스와 텍스트박스를 함께 처리)
        private void LoadSettingItem(CheckBox checkBox, TextBox textBox, CellSettingItem setting)
        {
            if (setting.Checked && !string.IsNullOrEmpty(setting.Value))
            {
                checkBox.IsChecked = true;
                textBox.IsEnabled = true;
                textBox.Text = setting.Value;
            }
            else if (setting.Checked)
            {
                checkBox.IsChecked = true;
                textBox.IsEnabled = true;
                textBox.Text = "";
            }
            else if (!setting.Checked && !string.IsNullOrEmpty(setting.Value))
            {
                checkBox.IsChecked = false;
                textBox.IsEnabled = false;
                textBox.Text = setting.Value;
            }
            else
            {
                checkBox.IsChecked = false;
                textBox.IsEnabled = false;
                textBox.Text = "";
            }
        }

        private void RegisterTextBoxEvents()
        {
            var textBoxes = new TextBox[]
            {
                txtLotNo, txtModelID, txtBuyerArticleNo, txtArticleID,
                txtInspectDate, txtName, txtProcessID, txtMachineID,
                txtInspectLevel, txtIRELevel, txtCustomID, txtInOutDate,
                txtFMLGubun, txtSumInspectQty, txtDefectYN, txtSumDefectQty
            };

            foreach (var textBox in textBoxes)
            {
                textBox.PreviewTextInput += TextBox_PreviewTextInput;
                textBox.PreviewKeyDown += TextBox_PreviewKeyDown;
            }
        }

        private void TextBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            foreach (char c in e.Text)
            {
                if (!((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || (c >= '0' && c <= '9')))
                {
                    lib.ShowTooltipMessage(sender as FrameworkElement, "특수 문자는 입력 할 수 없습니다.", MessageBoxImage.Stop, PlacementMode.Bottom);
                    e.Handled = true;
                    return;
                }
            }
        }

        private void TextBox_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.ImeProcessedKey != Key.HangulMode && e.Key == Key.ImeProcessed)
            {
                lib.ShowTooltipMessage(sender as FrameworkElement, "한글은 입력할 수 없습니다.", MessageBoxImage.Stop, PlacementMode.Bottom);
                e.Handled = true;
            }
        }

        private void CommonControl_Click(object sender, MouseButtonEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }

        private void CommonControl_Click(object sender, RoutedEventArgs e)
        {
            lib.CommonControl_Click(sender, e);
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        // 설정저장 버튼 (현재 설정 + 선택된 프리셋도 업데이트)
        private void btnSettingSave_Click(object sender, RoutedEventArgs e)
        {
            string currentText = cboPresets.Text?.Trim();
            bool isNewPreset = false;

            // 현재 텍스트가 기존 프리셋 이름과 다른지 확인
            if (currentSelectedPreset == null || currentSelectedPreset.Name != currentText)
            {
                // 새로운 이름이 입력된 경우
                if (!string.IsNullOrEmpty(currentText))
                {
                    isNewPreset = true;
                }
            }

            string message;
            if (isNewPreset)
            {
                // 새 프리셋 생성하는 경우
                message = $"'{currentText}' 새 프리셋을 저장하고 바로 적용하시겠습니까?";
            }
            else if (currentSelectedPreset != null)
            {
                // 기존 프리셋 업데이트하는 경우
                message = $"'{currentSelectedPreset.Name}' 프리셋으로 적용 및 설정저장 하시겠습니까?";
            }
            else
            {
                // 프리셋이 선택되지 않은 경우
                message = "설정을 저장 하시겠습니까?";
            }

            MessageBoxResult msgResult = MessageBox.Show(message, "확인", MessageBoxButton.YesNo);
            if (msgResult == MessageBoxResult.Yes)
            {
                if (isNewPreset)
                {
                    // 새 프리셋 생성 및 적용
                    CreateAndApplyNewPreset(currentText);
                }
                else if (currentSelectedPreset != null)
                {
                    // 기존 프리셋 업데이트
                    currentSelectedPreset.Settings = GetCurrentSettings();
                    currentSelectedPreset.CreatedDate = DateTime.Now;

                    MessageBox.Show($"'{currentSelectedPreset.Name}' 프리셋으로 설정이 저장되었습니다.");
                }
                else
                {
                    MessageBox.Show("설정이 저장되었습니다.");
                }

                SavePresetsAndSettings();
            }
            else
            {
                // 사용자가 '아니오'를 선택한 경우, 이전 선택된 프리셋으로 되돌리기
                if (isNewPreset && currentSelectedPreset != null)
                {
                    cboPresets.Text = currentSelectedPreset.Name;
                    cboPresets.SelectedItem = currentSelectedPreset;
                }
            }
        }

        // 새 프리셋 생성 및 적용 헬퍼 메서드
        private void CreateAndApplyNewPreset(string presetName)
        {
            // 이름 중복 확인
            var existing = presets.FirstOrDefault(p => p.Name == presetName);
            if (existing != null)
            {
                if (MessageBox.Show($"'{presetName}' 프리셋이 이미 존재합니다. 덮어쓰시겠습니까?",
                                   "확인", MessageBoxButton.YesNo) == MessageBoxResult.Yes)
                {
                    // 기존 프리셋 덮어쓰기
                    existing.Settings = GetCurrentSettings();
                    existing.CreatedDate = DateTime.Now;
                    currentSelectedPreset = existing;

                    RefreshPresetComboBox();
                    cboPresets.SelectedItem = existing;

                    MessageBox.Show($"'{presetName}' 프리셋이 업데이트되어 적용되었습니다.");
                }
                else
                {
                    // 덮어쓰기 거부 시 이전 프리셋으로 되돌리기
                    if (currentSelectedPreset != null)
                    {
                        cboPresets.Text = currentSelectedPreset.Name;
                        cboPresets.SelectedItem = currentSelectedPreset;
                    }
                }
            }
            else
            {
                // 새 프리셋 생성
                var newPreset = new CellSettingPreset
                {
                    Name = presetName,
                    Settings = GetCurrentSettings()
                };

                presets.Add(newPreset);
                currentSelectedPreset = newPreset;

                RefreshPresetComboBox();
                cboPresets.SelectedItem = newPreset;

                MessageBox.Show($"'{presetName}' 새 프리셋이 생성되어 적용되었습니다.");
            }
        }

        #endregion
    }

    #region 클래스 정의들

    // 개별 프리셋 클래스
    public class CellSettingPreset
    {
        public string Name { get; set; } = "새 프리셋";
        public CellSettings Settings { get; set; } = new CellSettings();
        public DateTime CreatedDate { get; set; } = DateTime.Now;

        public override string ToString()
        {
            return Name;
        }
    }

    // 전체 프리셋 컬렉션 클래스
    public class PresetCollection
    {
        public CellSettings CurrentSettings { get; set; } = new CellSettings();
        public List<CellSettingPreset> Presets { get; set; } = new List<CellSettingPreset>();
        public string LastSelectedPresetName { get; set; } = ""; // 마지막 선택된 프리셋 이름 저장
    }

    // 개별 설정 항목 클래스
    public class CellSettingItem
    {
        public bool Checked { get; set; } = false;
        public string Value { get; set; } = "";
    }

    // 전체 설정 클래스
    public class CellSettings
    {
        public CellSettingItem LotNo { get; set; } = new CellSettingItem();
        public CellSettingItem ModelID { get; set; } = new CellSettingItem();
        public CellSettingItem BuyerArticleNo { get; set; } = new CellSettingItem();
        public CellSettingItem ArticleID { get; set; } = new CellSettingItem();
        public CellSettingItem InspectDate { get; set; } = new CellSettingItem();
        public CellSettingItem Name { get; set; } = new CellSettingItem();
        public CellSettingItem ProcessID { get; set; } = new CellSettingItem();
        public CellSettingItem MachineID { get; set; } = new CellSettingItem();
        public CellSettingItem InspectLevel { get; set; } = new CellSettingItem();
        public CellSettingItem IRELevel { get; set; } = new CellSettingItem();
        public CellSettingItem CustomID { get; set; } = new CellSettingItem();
        public CellSettingItem InOutDate { get; set; } = new CellSettingItem();
        public CellSettingItem FMLGubun { get; set; } = new CellSettingItem();
        public CellSettingItem SumInspectQty { get; set; } = new CellSettingItem();
        public CellSettingItem DefectYN { get; set; } = new CellSettingItem();
        public CellSettingItem SumDefectQty { get; set; } = new CellSettingItem();
    }

    #endregion
}