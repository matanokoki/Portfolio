from PIL import Image
import numpy as np
import serial
import time
import sys

# --- 1. テスト用ビット列を作成 ---
data = np.array([1, 0, 1, 0, 1, 1, 0, 0, 1, 1, 1, 0, 0, 0, 1, 0,1,0,1,0,1,1,0,0,1,1,1,0,0,0,1,0,1,0,1,1,0,0,1,0,1,1,1,0,1,0,0,0,1,1])

print(f"テスト用ビット列: {data}")
print(f"総ビット数: {len(data)}")

symbol_string = ''.join(map(str, data))

# --- 4. エラーレート計算用の正解ファイル (ground_truth.txt) を作成 ---
output_filename = "ground_truth.txt"
print(f"正解データファイル '{output_filename}' を作成中...")
with open(output_filename, "w") as f:
    for bit in data: # プリアンブルを含まない元のデータ
        f.write(f"{bit}\n")
print(f"'{output_filename}' に {len(data)} ビット書き込みました。")


# --- 5. 同期のためのプリアンブル設定 ---
# 2-PAM用プリアンブル (例: 101011)
# 必ず信号の変化 (0と1の遷移) を含むパターンにする
preamble = "101011"
data_with_preamble = preamble + symbol_string # 送信するのはプリアンブル付き

print(f"プリアンブル ({preamble}) を追加しました。")
print(f"送信シンボル列 (一部): {data_with_preamble[:50]}...") # 最初の50文字だけ表示


# --- 6. Arduinoへシリアル送信 ---
# 一気にArduinoに送信ビット列を送信

try:
    # COMポートはご自身の環境に合わせて変更してください (例: 'COM3' や '/dev/ttyUSB0')
    SERIAL_PORT = 'COM3' 
    BAUD_RATE = 9600
    
    ser = serial.Serial(SERIAL_PORT, BAUD_RATE, timeout=1) 
    print(f"シリアルポート '{SERIAL_PORT}' を開きました。Arduinoのリセット待ち...")
    time.sleep(2)  # Arduinoのリセット・接続待ち

    print(f"Arduinoへデータを一括送信")

    ser.write(data_with_preamble.encode('ascii') + b'\n')
    
    ser.close()
    
    sending_time = len(data_with_preamble) * 0.066

    print("\nデータ送信完了。")
    print(f"Arduino点滅時間: {sending_time} 秒")
    print(f"ポート '{SERIAL_PORT}' を閉じました。")

except serial.SerialException as e:
    print(f"\nエラー: シリアルポート '{SERIAL_PORT}' が見つからないか、開けません。")
    print("Arduinoが接続されているか、COMポート番号が正しいか確認してください。")
    print(f"詳細: {e}")
except Exception as e:
    print(f"予期せぬエラーが発生しました: {e}")
    if 'ser' in locals() and ser.is_open:
        ser.close()
        print(f"ポート '{SERIAL_PORT}' を緊急で閉じました。")