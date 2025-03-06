import os
import sys
import comtypes.client
import csv

# SAP2000 uygulamasına bağlanma
AttachToInstance = False
SpecifyPath = False
ProgramPath = 'C:\Program Files\Computers and Structures\SAP2000 25\SAP2000.exe'
APIPath = input('Proje dosyalarının kaydedileceği klasör yolunu girin: ')
if not os.path.exists(APIPath):
    os.makedirs(APIPath)
ModelPath = APIPath + os.sep + 'SimpleBeam.sdb'

# CSV dosya yolu
csv_file_path = APIPath + os.sep + 'Simple_Beam_Results.csv'

# API yardımcı nesnesi oluşturma
helper = comtypes.client.CreateObject('SAP2000v1.Helper')
helper = helper.QueryInterface(comtypes.gen.SAP2000v1.cHelper)

if AttachToInstance:
    try:
        mySapObject = helper.GetObject('CSI.SAP2000.API.SapObject')
    except (OSError, comtypes.COMError):
        print('SAP2000 çalışmıyor veya bağlanılamadı.')
        sys.exit(-1)
else:
    if SpecifyPath:
        mySapObject = helper.CreateObject(ProgramPath)
    else:
        mySapObject = helper.CreateObjectProgID('CSI.SAP2000.API.SapObject')
    mySapObject.ApplicationStart()

# SAP2000 model nesnesini oluşturma
SapModel = mySapObject.SapModel
SapModel.InitializeNewModel()
SapModel.File.NewBlank()

# Birimleri metre-kilonewton olarak ayarlama
kN_m_C = 6
SapModel.SetPresentUnits(kN_m_C)

# 2 açıklıklı basit kirişin geometri tanımları
FrameName1 = ''
FrameName2 = ''

# İlk açıklık: 4 metre
[FrameName1, ret] = SapModel.FrameObj.AddByCoord(0, 0, 0, 4, 0, 0, FrameName1, 'Default', '1', 'Global')

# İkinci açıklık: 5 metre
[FrameName2, ret] = SapModel.FrameObj.AddByCoord(4, 0, 0, 9, 0, 0, FrameName2, 'Default', '2', 'Global')

# Görünümü yenileme
SapModel.View.RefreshView(0, False)

# Malzeme tanımlama (E=3x10^7, weight=0, adı='malzeme')
MATERIAL_NAME = 'malzeme'
MATERIAL_CONCRETE = 2  # 2: Concrete
SapModel.PropMaterial.SetMaterial(MATERIAL_NAME, MATERIAL_CONCRETE)
SapModel.PropMaterial.SetMPIsotropic(MATERIAL_NAME, 3e7, 0, 0)
SapModel.PropMaterial.SetWeightAndMass(MATERIAL_NAME, 1, 0)

# Sabit mesnetler (sadece dönme serbestliği)
SabitMesnet = [True, True, True, False, False, False]

# (0,0,0), (4,0,0) ve (9,0,0) noktalarına sabit mesnet ekleme
SapModel.PointObj.SetRestraint('1', SabitMesnet)
SapModel.PointObj.SetRestraint('2', SabitMesnet)
SapModel.PointObj.SetRestraint('3', SabitMesnet)

# (0,0,0) ve (4,0,0) arasında kalan çubuğa Z yönünde 20 kN/m yayılı yük tanımlama
SapModel.FrameObj.SetLoadDistributed(FrameName1, 'DEAD', 1, 10, 0, 1, 20, 20, 'Global')

# (6,0,0) noktasında yeni bir nokta oluşturma
PointName = ''
[PointName, ret] = SapModel.PointObj.AddCartesian(6, 0, 0)

# (6,0,0) noktasına aşağı yönlü -18 kN noktasal yük tanımlama
SapModel.PointObj.SetLoadForce(PointName, 'DEAD', [0, 0, -18, 0, 0, 0])
  
  

# SAP2000 modelini yeni isimle kaydetme
SapModel.File.Save(ModelPath)


# Analizi başlatma (XZ düzleminde)
SapModel.Analyze.RunAnalysis()

# Yükleme durumunu seçme
SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()
SapModel.Results.Setup.SetCaseSelectedForOutput('DEAD')

# Diğer noktalar için mesnet reaksiyonlarını alma
for point_name in ['1', '2', '3']:
    NumberResults, Obj, Elem, LoadCase, StepType, StepNum, F1, F2, F3, M1, M2, M3, ret = SapModel.Results.JointReact(point_name, 0, 0, [], [], [], [], [], [], [], [], [], [])
    if NumberResults > 0:
        print(f'Mesnet Reaksiyonları (Nokta {point_name}): Fx={F1}, Fy={F2}, Fz={F3}, Mx={M1}, My={M2}, Mz={M3}')
    else:
        print(f'Mesnet reaksiyonları bulunamadı (Nokta {point_name}). Lütfen analiz ve yük durumunu kontrol edin.')



# M, N, T (Moment, Normal kuvvet, Kesme kuvveti) sonuçlarını alma
for frame_name in [FrameName1, FrameName2]:
    NumberResults = 0
    Obj = []
    ObjSta = []
    Elm = []
    ElmSta = []
    LoadCase = []
    StepType = []
    StepNum = []
    P = []
    V2 = []
    V3 = []
    T = []
    M2 = []
    M3 = []

    [NumberResults, Obj, ObjSta, Elm, ElmSta, LoadCase, StepType, StepNum, P, V2, V3, T, M2, M3, ret] = SapModel.Results.FrameForce(
        frame_name, 0, 0, Obj, ObjSta, Elm, ElmSta, LoadCase, StepType, StepNum, P, V2, V3, T, M2, M3)

    if NumberResults > 0:
        for i in range(NumberResults):
            print(f'Çubuk {frame_name} için: Normal Kuvvet (N)={P[i]}, Kesme Kuvveti (T)={V2[i]}, Moment (M)={M3[i]}')
    else:
        print(f'Çubuk {frame_name} için sonuçlar bulunamadı. Lütfen analiz ve yük durumunu kontrol edin.')


# Sonuçları CSV dosyasına yazma
with open(csv_file_path, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)
    writer.writerow(['Frame Name', 'Normal Kuvvet (N)', 'Kesme Kuvveti (T)', 'Moment (M)'])

    # M, N, T (Moment, Normal kuvvet, Kesme kuvveti) sonuçlarını alma
    for frame_name in [FrameName1, FrameName2]:
        NumberResults = 0
        Obj = []
        ObjSta = []
        Elm = []
        ElmSta = []
        LoadCase = []
        StepType = []
        StepNum = []
        P = []
        V2 = []
        V3 = []
        T = []
        M2 = []
        M3 = []

        [NumberResults, Obj, ObjSta, Elm, ElmSta, LoadCase, StepType, StepNum, P, V2, V3, T, M2, M3, ret] = SapModel.Results.FrameForce(
            frame_name, 0, 0, Obj, ObjSta, Elm, ElmSta, LoadCase, StepType, StepNum, P, V2, V3, T, M2, M3)

        if NumberResults > 0:
            for i in range(NumberResults):
                writer.writerow([frame_name, P[i], V2[i], M3[i]])
                print(f'Çubuk {frame_name} için: Normal Kuvvet (N)={P[i]}, Kesme Kuvveti (T)={V2[i]}, Moment (M)={M3[i]}')
        else:
            print(f'Çubuk {frame_name} için sonuçlar bulunamadı. Lütfen analiz ve yük durumunu kontrol edin.')

print(f"Sonuçlar başarıyla '{csv_file_path}' dosyasına kaydedildi!")
