import cv2
import numpy as np
import matplotlib.pyplot as plt
import sys

# --- 定数 ---
PREAMBLE = np.array([1, 0, 1, 0, 1, 1]) # 送信側 (image_to_bit_2pam.py) と合わせる
PREAMBLE_INV = np.array([0, 1, 0, 1, 0, 0]) # PREAMBLEの反転
PREAMBLE_LEN = len(PREAMBLE)

def extract_signal_from_video(video_path):
    """
    (ステップ1) 動画を読み込み、ROIの平均輝度から信号を抽出する
    """
    print("--- ステップ1: 動画から信号を抽出 ---")
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        print(f"エラー: 動画ファイル '{video_path}' を開けません。")
        return None

    ret, frame = cap.read()
    if not ret:
        print("エラー: 動画から最初のフレームを読み込めません。")
        cap.release()
        return None

    frame_gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY) 

    print("--- ROI選択 ---")
    print("LEDの領域をマウスでドラッグして囲ってください。")
    print("領域を選択したら [Enter] キーを、キャンセルは [c] キーを押してください。")
    
    cv2.namedWindow("Select ROI", cv2.WINDOW_NORMAL)
    cv2.resizeWindow("Select ROI", 800, 600)
    roi = cv2.selectROI("Select ROI", frame, fromCenter=False, showCrosshair=True)
    cv2.destroyWindow("Select ROI")
    
    if roi[2] == 0 or roi[3] == 0:
        print("ROIが選択されませんでした。処理を中断します。")
        cap.release()
        return None

    x, y, w, h = roi
    print(f"ROIが選択されました: x={x}, y={y}, w={w}, h={h}")

    signal = []
    frame_count = 0
    
    roi_gray = frame_gray[y:y+h, x:x+w]
    signal.append(np.mean(roi_gray))
    frame_count += 1

    while True:
        ret, frame = cap.read()
        if not ret:
            break
        
        frame_gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY) 
        roi_gray = frame_gray[y:y+h, x:x+w]
        
        signal.append(np.mean(roi_gray))
        frame_count += 1

    cap.release()
    cv2.destroyAllWindows()
    
    print(f"動画全 {frame_count} フレームから信号を抽出完了。")
    return np.array(signal)


def demodulate_signal(rx_signal, user_symbol_rate_guess):
    """
    (ステップ2) 論文[2]の手法（微分と相互相関）で同期・復調する
    """
    print("\n--- ステップ2: 信号の同期と復調 ---")
    
    rx_signal_norm = (rx_signal - np.min(rx_signal)) / (np.max(rx_signal) - np.min(rx_signal))
    rx_diff = np.gradient(rx_signal_norm)
    rx_diff_abs = np.abs(rx_diff)

    best_phase = 0
    best_rate = 0
    max_score = -1

    rate_min = round(user_symbol_rate_guess * 0.7) 
    rate_max = round(user_symbol_rate_guess * 1.3)
    
    if rate_min < 1: rate_min = 1
    if rate_max < rate_min: rate_max = rate_min 

    print(f"シンボルレートを {rate_min}〜{rate_max} フレーム/sym の範囲で探索します...")

    for rate_guess in range(rate_min, rate_max + 1):
        if rate_guess == 0: continue
        for phase in range(rate_guess):
            samples = rx_diff_abs[phase::rate_guess]
            
            # --- ★★★ 重大なバグ修正 ★★★ ---
            # np.sum() はサンプル数が多い(rate=1)方が有利になるため
            # np.mean() (平均) で公平に比較する
            score = np.mean(samples) 
            
            if score > max_score:
                max_score = score
                best_phase = phase
                best_rate = rate_guess

    if best_rate == 0 or max_score <= 0: # max_scoreが0以下の場合もエラーとする
        print(f"エラー: 同期信号を検出できませんでした。(スコア: {max_score})")
        return None, None
        
    print(f"同期完了: 検出されたシンボルレート = {best_rate:.2f} フレーム/sym")
    print(f"検出された遷移位相 (シンボル境界) = {best_phase} フレーム目")

    sampling_phase = (best_phase + (best_rate // 2)) % best_rate
    print(f"サンプリング位相 (シンボル中央) = {sampling_phase} フレーム目")

    sampled_signal = rx_signal_norm[sampling_phase::best_rate]
    rx_bits_all = (sampled_signal > 0.5).astype(int)

    data_start_index = -1
    is_inverted = False
    
    for i in range(len(rx_bits_all) - PREAMBLE_LEN):
        window = rx_bits_all[i : i + PREAMBLE_LEN]
        if np.array_equal(window, PREAMBLE):
            data_start_index = i + PREAMBLE_LEN
            print(f"プリアンブル {PREAMBLE} をインデックス {i} で検出。")
            break
            
    if data_start_index == -1:
        print("プリアンブルが見つかりません。反転信号を試します...")
        for i in range(len(rx_bits_all) - PREAMBLE_LEN):
            window = rx_bits_all[i : i + PREAMBLE_LEN]
            if np.array_equal(window, PREAMBLE_INV):
                data_start_index = i + PREAMBLE_LEN
                is_inverted = True
                print(f"反転プリアンブル {PREAMBLE_INV} をインデックス {i} で検出。")
                break

    if data_start_index == -1:
        print("エラー: プリアンブルも反転プリアンブルも検出できませんでした。同期に失敗しました。")
        return None, None

    rx_bits_payload = rx_bits_all[data_start_index:]
    
    if is_inverted:
        print("信号が反転していると判断し、ビットを反転します。")
        rx_bits_payload = 1 - rx_bits_payload 

    print(f"復調（ペイロード）完了。 {len(rx_bits_payload)} ビットを検出。")

    plot_data = {
        "raw_signal": rx_signal_norm,
        "transition_phase": best_phase,
        "sampling_phase": sampling_phase,
        "rate": best_rate,
        "sampled_values": sampled_signal,
        "rx_bits": (sampled_signal > 0.5).astype(int) 
    }
    
    return rx_bits_payload, plot_data


def evaluate_ber(sent_symbols_path, rx_bits):
    """
    (ステップ3) 正解データと復調結果を比較し、BERを計算する
    """
    print("\n--- Step3: Evaluate error rate ---")
    try:
        sent_bits = np.loadtxt(sent_symbols_path, dtype=int, comments="#")
        if sent_bits.ndim == 0:
            sent_bits = np.array([sent_bits])
        print(f"Loaded '{sent_symbols_path}' (True bit: {len(sent_bits)} bit)")
        
    except FileNotFoundError:
        print(f"エラー: 正解データファイル '{sent_symbols_path}' が見つかりません。")
        return
    except Exception as e:
        print(f"エラー: 正解データファイル '{sent_symbols_path}' の読み込み中にエラーが発生しました。: {e}")
        return

    n_bits_to_compare = min(len(sent_bits), len(rx_bits))
    
    if n_bits_to_compare == 0:
        print("エラー: 比較できるビットがありません。")
        return

    print(f"Compare the first {n_bits_to_compare} bits of the correct answer data and the received data.")
    
    sent_bits_truncated = sent_bits[:n_bits_to_compare]
    rx_bits_truncated = rx_bits[:n_bits_to_compare]

    bit_errors = np.sum(sent_bits_truncated != rx_bits_truncated)
    total_bits = n_bits_to_compare
    ber = bit_errors / total_bits

    print("\n--- Bit Error Rate(BER) ---")
    print(f"Comparison Bit Size: {total_bits}")
    print(f"Bit error count: {bit_errors}")
    print(f"Bit error rate: {ber * 100:.2f} %")
    
    print(f"True: {sent_bits_truncated}")
    print(f"Received: {rx_bits_truncated}")


def plot_results(plot_data):
    """
    (ステップ4) 復調結果をグラフにプロットする
    """
    
    raw_signal = plot_data["raw_signal"]
    transition_phase = plot_data["transition_phase"]
    sampling_phase = plot_data["sampling_phase"]
    rate = plot_data["rate"]
    sampled_values = plot_data["sampled_values"]
    rx_bits = plot_data["rx_bits"]

    plt.figure(figsize=(15, 6))
    
    plt.plot(raw_signal, label="Raw Signal (Normalized)", color='blue', alpha=0.6)
    
    for i in range(transition_phase, len(raw_signal), rate):
        plt.axvline(x=i, color='gray', linestyle='--', linewidth=0.8, label="Symbol Boundary (Detection)" if i==transition_phase else "")
        
    sample_indices = np.arange(sampling_phase, len(raw_signal), rate)
    valid_indices = sample_indices[sample_indices < len(raw_signal)]
    
    actual_sampled_values = raw_signal[valid_indices]
    
    bits_for_plot = (actual_sampled_values > 0.5).astype(int)
    
    indices_0 = valid_indices[bits_for_plot == 0]
    values_0 = actual_sampled_values[bits_for_plot == 0]
    
    indices_1 = valid_indices[bits_for_plot == 1]
    values_1 = actual_sampled_values[bits_for_plot == 1]

    plt.plot(indices_0, values_0, 'mo', label='sampling point (→ 0)', markersize=5)
    plt.plot(indices_1, values_1, 'go', label='sampling point (→ 1)', markersize=5)

    plt.title("Synchronization and Sampling Results")
    plt.xlabel("video frames")
    plt.ylabel("Normalized luminance")
    plt.legend()
    plt.grid(True)
    # グラフの表示範囲を、検出されたレートに基づいて調整
    #plt.xlim(0, min(len(raw_signal), 50 * rate)) 
    plt.tight_layout()
    plt.show()


def main():
    """
    メイン実行関数
    """
    try:
        video_path = input("処理する動画ファイルのパスを入力してください (例: image.mp4): ")
        sent_symbols_path = input("正解データファイルのパスを入力してください (例: ground_truth.txt): ")
        user_symbol_rate_guess = float(input("1シンボルのおよそのフレーム数を入力してください (例: 1.2): "))
        
        raw_signal = extract_signal_from_video(video_path)
        
        if raw_signal is None:
            print("信号抽出に失敗しました。処理を終了します。")
            return

        rx_bits, plot_data = demodulate_signal(raw_signal, user_symbol_rate_guess)
        
        if rx_bits is None or plot_data is None:
            print("復調に失敗しました。処理を終了します。")
            return

        evaluate_ber(sent_symbols_path, rx_bits)
        plot_results(plot_data)

    except Exception as e:
        print(f"\n予期せぬエラーが発生しました: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()

