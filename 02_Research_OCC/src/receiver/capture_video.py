import cv2
import time
import os

# --- 設定項目 ---
CAMERA_INDEX = 0
FRAME_WIDTH = 640
FRAME_HEIGHT = 480
FPS = 30.0
# -----------------

def main():
    """
    カメラ映像を表示し、キー操作で録画の開始・停止・終了を行うメイン関数
    """
    cap = cv2.VideoCapture(CAMERA_INDEX)
    if not cap.isOpened():
        print("エラー: カメラを開けませんでした。")
        return

    # --- マニュアル設定 ---
    cap.set(cv2.CAP_PROP_FPS, FPS)
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, FRAME_WIDTH)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, FRAME_HEIGHT)
    cap.set(cv2.CAP_PROP_AUTOFOCUS, 0)
    cap.set(cv2.CAP_PROP_AUTO_EXPOSURE, 1)
    cap.set(cv2.CAP_PROP_EXPOSURE, -6)
    cap.set(cv2.CAP_PROP_GAIN, 200)

    print("\nカメラのマニュアル設定を試みました。")
    print("（注意: お使いのカメラによっては、これらの設定が反映されない場合があります）")

    fourcc = cv2.VideoWriter_fourcc(*'mp4v')
    out = None
    is_recording = False

    print("\nカメラの準備ができました。")
    print("----------------------------------------------------")
    print("操作方法:")
    print("  's' キー: 録画を開始 (Start)")
    print("  'p' キー: 録画を停止 (Stop)")
    print("  'q' キー: プログラムを終了 (Quit)")
    print("----------------------------------------------------")

    while True:
        ret, frame = cap.read()
        if not ret:
            print("エラー: フレームを読み取れませんでした。")
            break

        display_frame = frame.copy()

        if is_recording:
            cv2.circle(display_frame, (30, 30), 10, (0, 0, 255), -1)
            cv2.putText(display_frame, 'REC', (50, 38), cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 255), 2)

        cv2.imshow('Camera Feed', display_frame)
        key = cv2.waitKey(1) & 0xFF

        if key == ord('s') and not is_recording:
            filename =  "50bit_150cm.mp4"
            
            # ★★★ 修正点①: 実際に取得したフレームのサイズを使用する ★★★
            # これで、設定したサイズと実際のサイズが違う場合の問題を回避できます。
            actual_height, actual_width, _ = frame.shape
            actual_size = (actual_width, actual_height)
            
            out = cv2.VideoWriter(filename, fourcc, FPS, actual_size)

            # ★★★ 修正点②: VideoWriterが正常に開けたか必ずチェックする ★★★
            if not out.isOpened():
                print("エラー: VideoWriterの初期化に失敗しました。")
                print("コーデックがシステムにインストールされているか、確認してください。")
                is_recording = False
                out = None
            else:
                is_recording = True
                print(f"[録画開始] -> 保存先: {filename} (サイズ: {actual_width}x{actual_height})")

        elif key == ord('p') and is_recording:
            is_recording = False
            if out is not None:
                out.release()
                out = None
            print("[録画停止]")
        
        elif key == ord('q'):
            print("プログラムを終了します。")
            break

        if is_recording and out is not None:
            out.write(frame)

    if out is not None:
        out.release()
    cap.release()
    cv2.destroyAllWindows()
    print("リソースを解放しました。")

if __name__ == '__main__':
    main()