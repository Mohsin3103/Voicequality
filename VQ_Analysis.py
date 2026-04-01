from nisqa.NISQA_model import nisqaModel
import os
import pandas as pd

def analyze_audio(wav_path, call_id, signalr_connection=None):
    """
    Analyze audio file using NISQA and send results to SignalR
    
    Args:
        wav_path: Path to the WAV file
        call_id: Call ID (VOICE_ID)
        signalr_connection: SignalR connection object
    
    Returns:
        tuple: (mos, noise, distortion, loudness, verdict)
    """
    try:
        args = {
            'mode': 'predict_file',
            'pretrained_model': r'C:\NISQA\weights\nisqa.tar',
            'deg': wav_path,
            'output_dir': r'C:\Users\eommhoh\Desktop\VOICERecordings',
            'tr_bs_val': 1,
            'tr_num_workers': 0,
            'ms_channel': 0,
            'ms_max_segments': 10000,
        }
        
        print(f"🔍 Loading NISQA model for {call_id}...")
        model = nisqaModel(args)
        
        print(f"🎧 Running prediction for {call_id}...")
        results = model.predict()
        
        if isinstance(results, pd.DataFrame) and not results.empty:
            mos = float(results['mos_pred'][0])
            noise = float(results['noi_pred'][0])
            distortion = float(results['dis_pred'][0])
            loudness = float(results['loud_pred'][0])
            
            if mos > 2 and noise > 2.5 and distortion > 2 and loudness > 2.5:
                verdict = "GOOD"
            else:
                verdict = "POOR"
            
            print(f"\n🔈 QUALITY SUMMARY for {call_id}:")
            print(f"🎯 MOS: {mos:.2f}")
            print(f"🔉 Noise: {noise:.2f}")
            print(f"🎚️ Distortion: {distortion:.2f}")
            print(f"🔊 Loudness: {loudness:.2f}")
            print(f"➡️ Verdict: {verdict}")
            
            # Send results to SignalR
            if signalr_connection:
                try:
                    signalr_connection.send("SendAnalysisResult", [{
                        "callId": call_id,
                        "MOS": str(mos)
                    }])
                    
                    signalr_connection.send("SendAnalysisResult", [{
                        "callId": call_id,
                        "Noise": str(noise)
                    }])
                    
                    signalr_connection.send("SendAnalysisResult", [{
                        "callId": call_id,
                        "Distortion": str(distortion)
                    }])
                    
                    signalr_connection.send("SendAnalysisResult", [{
                        "callId": call_id,
                        "Loudness": str(loudness)
                    }])
                    
                    signalr_connection.send("SendAnalysisResult", [{
                        "callId": call_id,
                        "Verdict": verdict
                    }])
                    
                    print(f"✅ Analysis results sent to SignalR for {call_id}")
                except Exception as e:
                    print(f"❌ Failed to send analysis results to SignalR: {e}")
            
            return mos, noise, distortion, loudness, verdict
        else:
            print(f"❌ No prediction results for {call_id}")
            return None
            
    except Exception as e:
        print(f"❌ NISQA analysis failed for {call_id}: {e}")
        return None
