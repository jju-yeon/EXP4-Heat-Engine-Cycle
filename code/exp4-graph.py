import matplotlib.pyplot as plt
from matplotlib import font_manager, rc
import pandas as pd
import math
import numpy as np


#한글 폰트 설정
import matplotlib as mpl
mpl.rcParams['axes.unicode_minus'] = False
font_path = "C:/Windows/Fonts/malgun.ttf"
font_name = font_manager.FontProperties(fname=font_path).get_name()
rc('font', family=font_name)

#실험 이름 설정
expName = "EXP4-3"
expTheme = "사이클3"

#상수 설정
chamber_radius = 0.02
chamber_height = 0.082
cylinder_radius = 0.01625
startHeight = 0.019
sensor_radius = 0.0145

m_weight = 0.2
g = 9.81


#엑셀 불러오기
df = pd.read_excel("Capstone_Data-EXP4.xlsx", engine="openpyxl")

#불러올 행 결정하기
cols = [f"Absolute Pressure (kPa) {expTheme}",
        f"angle (rad) {expTheme}"]

#빈 셀 제거
sub = df[cols].dropna()

T = [f"Hot Temperature (1) (°C) {expTheme}",
     f"Cold Temperature (2) (°C) {expTheme}"]

T = df[T].dropna()

#P, Angle, T 값 불러오기
P = sub[f"Absolute Pressure (kPa) {expTheme}"].to_numpy()
Angle = sub[f"angle (rad) {expTheme}"].to_numpy()
T_high = T[f"Hot Temperature (1) (°C) {expTheme}"].to_numpy()
T_low = T[f"Cold Temperature (2) (°C) {expTheme}"].to_numpy()

#단위 변환
P *= 1E3
x = Angle * sensor_radius
V = (x + startHeight) * np.pi * cylinder_radius**2
T_high += 273.15
T_low += 273.15

T_H = T_high[0]
T_L = T_low[0]

#이론적 최대 열효율
eff_theory = (1-T_L/T_H)*100
print(f"이론적 최대 열효율: {eff_theory:.6f}%")

#한 일 구하기
P = np.append(P, P[0])
V = np.append(V, V[0])
W = np.trapezoid(P, V)
print(f"수치적분: {W:.6f}")

#포인트 지정
#상위 30%의 부피 데이터만 모으기
threshold = np.percentile(V, 70)
idx = np.where(V >= threshold)[0]
P_right = P[idx]
V_right = V[idx]
#상위 30%의 부피 데이터 중 가장 큰 압력과 낮은 압력 구하기
P_C = P_right[np.argmax(P_right)]
P_D = P_right[np.argmin(P_right)]


#부피 계산
V_A = np.pi * (chamber_radius**2) * chamber_height \
      + np.pi * (cylinder_radius**2) * startHeight
V_D = (T_H/T_L) * V_A
V_C = (P_D/P_C) * V_D

#열효율 계산
Q_CD = P_D * V_D * math.log(V_D/V_C)
Q_BC = (7/2) * P_C * V_C * (T_H-T_L)/T_H

Q_H = Q_BC + Q_CD

eff = W/Q_H * 100

print(f"열효율: {eff:.6f}%")
print("카르노 효율 비: %f%%"%((eff/eff_theory)*100))
#질량체가 받은 일
delta_H = V_right[np.argmax(P_right)]/(np.pi*cylinder_radius**2) - startHeight
W_weight = m_weight * g * delta_H
print(f"질량체가 받은 일: {W_weight:.6f}J")

#PVT Graph 그리기
plt.figure()
plt.scatter(V, P, s=5)
plt.ylabel("Pressur(Pa)")
plt.xlabel("Volume(m³)")
plt.title(f"PV Graph ({expName})")
          
plt.grid(True)
plt.savefig(f"PVT Graph ({expName})", dpi=300)


plt.show()
