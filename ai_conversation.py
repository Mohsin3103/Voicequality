"""
AI-Powered Conversational Module for Voice Quality Testing
TTS-first stable version for Jabber/VB-Cable call routing test
"""

import os
import time
import logging
import re
import random
import shutil
import tempfile
import pyttsx3
import speech_recognition as sr
import subprocess
import winsound
try:
    from openai import OpenAI
except ImportError:
    OpenAI = None
try:
    import win32com.client
except ImportError:
    win32com = None

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

VB_CABLE_INPUT_DEVICE = "CABLE Output (2- VB-Audio Virtual Cable)"
VB_CABLE_PLAYBACK_DEVICE = "Speakers (2- VB-Audio Virtual Cable)"
FFMPEG_TTS_CANDIDATES = [
    os.getenv("FFMPEG_PATH"),
    r"C:\Users\eommhoh\AppData\Local\Microsoft\WinGet\Packages\Gyan.FFmpeg_Microsoft.Winget.Source_8wekyb3d8bbwe\ffmpeg-8.0-full_build\bin\ffmpeg.exe",
    shutil.which("ffmpeg"),
]
MAX_DYNAMIC_RESPONSE_SECONDS = 5


class AIConversationAgent:
    def __init__(
        self,
        api_key=None,
        input_device_name=VB_CABLE_INPUT_DEVICE,
    ):
        self.api_key = api_key or os.getenv("OPENAI_API_KEY")
        self.conversation_history = []
        self.audio_quality = None
        self.cli_confirmed = None
        self.cli_number = None
        self.playback_routing_ready = False
        self.tts_cache = {}

        self.input_device_name = input_device_name
        self.ffmpeg_path = self._resolve_ffmpeg_path()
        self.flite_voice = "slt"
        self.flite_tempo = 0.92
        self.openai_tts_voice = "shimmer"
        self.openai_tts_model = "gpt-4o-mini-tts"
        self.openai_tts_speed = 0.86
        self.openai_tts_instructions = (
            "Speak as Monica, a warm, professional young woman. "
            "Use a clear, calm, welcoming tone with natural pacing for a phone call. "
            "Pronounce digits carefully and leave a short gap between them."
        )
        self.openai_client = None
        if self.api_key and OpenAI is not None:
            try:
                self.openai_client = OpenAI(api_key=self.api_key)
                logger.info("OpenAI transcription client ready.")
            except Exception as e:
                logger.warning(f"OpenAI client initialization failed: {e}")

        # -------- TTS engine --------
        self.tts_mode = "ffmpeg_flite"
        self.tts_engine = None
        self.sapi_voice = None
        self._init_tts()

        # -------- STT recognizer --------
        self.recognizer = sr.Recognizer()
        self.recognizer.energy_threshold = 180
        self.recognizer.dynamic_energy_threshold = True
        self.recognizer.dynamic_energy_adjustment_damping = 0.15
        self.recognizer.dynamic_energy_ratio = 1.5
        self.recognizer.pause_threshold = 0.8
        self.recognizer.phrase_threshold = 0.2
        self.recognizer.non_speaking_duration = 0.35

        self.mic_index = self._find_input_device(self.input_device_name)

        try:
            with sr.Microphone(device_index=self.mic_index) as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=0.2)
        except Exception as e:
            logger.warning(f"Initial mic calibration failed: {e}")

        self.system_prompt = """You are a friendly and professional voice quality testing assistant for Mobily, Saudi Arabia.

Your name is Monica. You are a Mobily Voice Quality Test Agent with a natural, welcoming female voice.
Your task is to follow this EXACT conversation flow:
1. Greet the user warmly and introduce yourself as Monica from Mobily
2. Inform them that the call is being recorded for quality purposes and ask for permission
3. If the user denies permission, politely end the call immediately
4. Ask if the user is free and available for a quick voice quality test
5. If the user is not available, politely end the call immediately
6. If the user is available, ask the user to confirm the caller ID number they see on screen
7. Ask for voice quality feedback and classify it as Excellent, Good, Average, Bad, or Poor
8. Thank the user warmly and end the call

Keep responses short, natural, and reassuring.
Follow the steps in ORDER.
"""

        logger.info(f"Mic index: {self.mic_index}")
        logger.info(f"Ensure Windows default playback is routed to {VB_CABLE_PLAYBACK_DEVICE}.")
        logger.info(f"Ensure Jabber microphone is set to {VB_CABLE_INPUT_DEVICE}")

    def _resolve_ffmpeg_path(self):
        for candidate in FFMPEG_TTS_CANDIDATES:
            if candidate and os.path.isfile(candidate):
                return candidate
        return None

    def _sanitize_flite_text(self, text):
        cleaned = re.sub(r"[^A-Za-z0-9\s]", " ", str(text or ""))
        cleaned = re.sub(r"\s+", " ", cleaned).strip()
        return cleaned or "Hello"

    def _prepare_tts_text(self, text):
        prepared = str(text or "")
        replacements = {
            "ITMC": "I T M C",
            "CLI": "caller I D",
            "Caller ID": "caller I D",
            "caller ID": "caller I D",
            "1-5": "1 to 5",
            "Hello, this is": "Hello, this is",
        }
        for src, dst in replacements.items():
            prepared = prepared.replace(src, dst)
        return self._sanitize_flite_text(prepared)

    def _cacheable_prompts(self):
        return [
            "Hello, this is Monica from Mobily Voice Quality Testing. It is so nice to speak with you today.",
            "This call is being recorded for quality purposes. Do I have your permission to continue?",
            "Thank you for confirming. Are you free and available for a quick voice quality test?",
            "Thank you very much for your time. Goodbye.",
        ]

    def _build_transcription_prompt(self, context, expected_length=None):
        prompts = {
            "otp": "Transcribe only the 4 digit one time password repeated by the caller. Return only the four spoken digits in order. If unclear, return the single word unclear.",
            "quality": "Transcribe only the caller answer. It will usually be one short rating like 1, 2, 3, 4, 5, one, two, three, four, five, poor, average, good, or excellent. If unclear, return the single word unclear.",
            "confirm": "Transcribe only a short confirmation answer like yes, yeah, correct, right, no, wrong, or incorrect. If unclear, return the single word unclear.",
            "cli": "Transcribe only the caller ID digits spoken by the caller. Preserve all spoken digits in order. If unclear, return the single word unclear.",
            "status": "Transcribe only a short wellbeing answer like good, fine, okay, not good, busy, tired, or bad. If unclear, return the single word unclear.",
            "free": "Transcribe only the caller speech clearly and briefly. If unclear, return the single word unclear.",
        }
        if context == "cli" and expected_length:
            return (
                f"Transcribe only the caller ID digits spoken by the caller. "
                f"The caller ID should contain exactly {expected_length} digits. "
                f"Preserve every spoken digit in order and return digits only. "
                f"If any digit is missing or unclear, return the single word unclear."
            )
        if context == "cli":
            return (
                "Transcribe only the caller ID digits spoken by the caller. "
                "The caller ID will usually be 9 or 10 digits long. "
                "Preserve every spoken digit in order and return digits only. "
                "If any digit is missing or unclear, return the single word unclear."
            )
        return prompts.get(context, prompts["free"])

    def _capture_response_audio(self, timeout, phrase_time_limit):
        temp_wav = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as tmp:
                temp_wav = tmp.name

            capture_seconds = min(
                max(
                    2.0,
                    float(phrase_time_limit if phrase_time_limit is not None else MAX_DYNAMIC_RESPONSE_SECONDS),
                ),
                float(MAX_DYNAMIC_RESPONSE_SECONDS),
            )

            if self.ffmpeg_path:
                subprocess.run(
                    [
                        self.ffmpeg_path,
                        "-y",
                        "-f",
                        "dshow",
                        "-i",
                        f"audio={self.input_device_name}",
                        "-t",
                        f"{capture_seconds:.2f}",
                        "-ac",
                        "1",
                        "-ar",
                        "16000",
                        temp_wav,
                    ],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=True,
                    timeout=max(10, int(capture_seconds) + 6),
                )
            else:
                with sr.Microphone(device_index=self.mic_index) as source:
                    self.recognizer.adjust_for_ambient_noise(source, duration=0.2)
                    audio = self.recognizer.listen(
                        source,
                        timeout=timeout,
                        phrase_time_limit=min(
                            phrase_time_limit if phrase_time_limit is not None else MAX_DYNAMIC_RESPONSE_SECONDS,
                            MAX_DYNAMIC_RESPONSE_SECONDS,
                        ),
                    )
                wav_bytes = audio.get_wav_data(convert_rate=16000, convert_width=2)
                if not wav_bytes:
                    return None
                with open(temp_wav, "wb") as wav_file:
                    wav_file.write(wav_bytes)

            if os.path.exists(temp_wav) and os.path.getsize(temp_wav) > 2048:
                return temp_wav
        except sr.WaitTimeoutError:
            return None
        except Exception as e:
            logger.warning(f"Dynamic caller audio capture failed: {e}")

        if temp_wav and os.path.exists(temp_wav):
            try:
                os.remove(temp_wav)
            except OSError:
                pass
        return None

    def _transcribe_with_openai(self, audio_path, context="free", expected_length=None):
        if not self.openai_client or not audio_path or not os.path.exists(audio_path):
            return None

        try:
            with open(audio_path, "rb") as audio_file:
                transcript = self.openai_client.audio.transcriptions.create(
                    model="gpt-4o-transcribe",
                    file=audio_file,
                    prompt=self._build_transcription_prompt(context, expected_length=expected_length),
                    language="en",
                    temperature=0,
                )

            text = (getattr(transcript, "text", "") or "").strip()
            if text:
                if text.lower().strip() in {"unclear", "unknown", "inaudible"}:
                    return None
                text = self._normalize_recognized_text(text, context=context)
                logger.info(f"Caller said (OpenAI): {text}")
                print(f"Caller: {text}")
                return text
        except Exception as e:
            logger.warning(f"OpenAI transcription failed for context={context}: {e}")
        return None

    def _preprocess_audio_for_transcription(self, audio_path):
        if not self.ffmpeg_path or not audio_path or not os.path.exists(audio_path):
            return audio_path

        processed_path = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as tmp:
                processed_path = tmp.name

            subprocess.run(
                [
                    self.ffmpeg_path,
                    "-y",
                    "-i",
                    audio_path,
                    "-af",
                    "highpass=f=120,lowpass=f=3400,afftdn,loudnorm",
                    "-ac",
                    "1",
                    "-ar",
                    "16000",
                    processed_path,
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=True,
                timeout=15,
            )
            return processed_path
        except Exception as e:
            logger.warning(f"Audio preprocessing failed: {e}")
            if processed_path and os.path.exists(processed_path):
                try:
                    os.remove(processed_path)
                except OSError:
                    pass
            return audio_path

    def _normalize_recognized_text(self, text, context="free"):
        normalized = (text or "").strip()
        lower = normalized.lower()

        if context == "status":
            if any(phrase in lower for phrase in ["good", "doing good", "doing well", "fine", "great", "well", "okay", "ok"]):
                return "I am good"
            if any(phrase in lower for phrase in ["not good", "bad", "tired", "busy", "sad", "sick"]):
                return "I am not good"

        if context == "confirm":
            if any(word in lower for word in ["yes", "yeah", "correct", "right", "confirmed", "true", "exactly"]):
                return "yes"
            if any(word in lower for word in ["no", "wrong", "incorrect", "negative", "not correct"]):
                return "no"

        if context == "quality":
            quality_map = {
                "1": "1",
                "2": "2",
                "3": "3",
                "4": "4",
                "5": "5",
                "one": "1",
                "two": "2",
                "three": "3",
                "four": "4",
                "five": "5",
                "won": "1",
                "to": "2",
                "too": "2",
                "tree": "3",
                "for": "4",
                "poor": "1",
                "bad": "1",
                "average": "3",
                "good": "4",
                "excellent": "5",
            }
            for src, dst in quality_map.items():
                if re.search(rf"\b{re.escape(src)}\b", lower):
                    return dst

        if context == "cli":
            return self._normalize_digits_from_speech(normalized) or normalized

        if context == "otp":
            digits = self._normalize_digits_from_speech(normalized)
            return digits if len(digits) == 4 else normalized

        return normalized

    def _init_tts(self):
        if self.openai_client is not None:
            self.tts_mode = "openai_tts"
            logger.info(f"Using OpenAI TTS voice for call audio: {self.openai_tts_voice}")
            return

        if self.ffmpeg_path:
            self.tts_mode = "ffmpeg_flite"
            logger.info(f"Using FFmpeg flite voice for call audio: {self.ffmpeg_path}")
            return

        if win32com is not None:
            try:
                self.sapi_voice = win32com.client.Dispatch("SAPI.SpVoice")
                try:
                    female_voice = None
                    for token in self.sapi_voice.GetVoices():
                        desc = (token.GetDescription() or "").lower()
                        if any(name in desc for name in ["zira", "hazel", "aria", "female"]):
                            female_voice = token
                            break
                    if female_voice is not None:
                        self.sapi_voice.Voice = female_voice
                        logger.info(f"Selected female SAPI voice: {female_voice.GetDescription()}")
                except Exception as e:
                    logger.warning(f"SAPI female voice selection failed: {e}")
                self.sapi_voice.Rate = 0
                self.sapi_voice.Volume = 100
                self.tts_mode = "sapi"
                logger.info("Using native Windows SAPI voice for call audio.")
                return
            except Exception as e:
                logger.warning(f"SAPI voice initialization failed: {e}")

        self.tts_mode = "pyttsx3"
        try:
            self.tts_engine = pyttsx3.init()
            self.tts_engine.setProperty("rate", 155)
            self.tts_engine.setProperty("volume", 1.0)

            try:
                voices = self.tts_engine.getProperty("voices")
                for voice in voices:
                    name = (getattr(voice, "name", "") or "").lower()
                    if "female" in name or "zira" in name:
                        self.tts_engine.setProperty("voice", voice.id)
                        logger.info(f"Selected pyttsx3 voice: {getattr(voice, 'name', voice.id)}")
                        break
            except Exception as e:
                logger.warning(f"Voice selection failed: {e}")

            self.tts_engine.say(" ")
            self.tts_engine.runAndWait()
            logger.info("Using pyttsx3 voice for call audio.")
            return
        except Exception as e:
            logger.warning(f"pyttsx3 initialization failed: {e}")
            self.tts_engine = None

    def _speak_with_openai_tts(self, text):
        if self.openai_client is None:
            raise RuntimeError("OpenAI TTS client is not available")

        temp_audio = None
        temp_wav = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".mp3", delete=False) as tmp:
                temp_audio = tmp.name
            with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as tmp:
                temp_wav = tmp.name

            response = self.openai_client.audio.speech.create(
                model=self.openai_tts_model,
                voice=self.openai_tts_voice,
                input=text,
                instructions=self.openai_tts_instructions,
                response_format="mp3",
                speed=self.openai_tts_speed,
            )
            response.write_to_file(temp_audio)

            if self.ffmpeg_path:
                subprocess.run(
                    [
                        self.ffmpeg_path,
                        "-y",
                        "-i",
                        temp_audio,
                        "-ac",
                        "1",
                        "-ar",
                        "16000",
                        temp_wav,
                    ],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=True,
                    timeout=20,
                )
            else:
                raise RuntimeError("FFmpeg is required for OpenAI TTS playback conversion")

            winsound.PlaySound(temp_wav, winsound.SND_FILENAME)
        finally:
            if temp_audio and os.path.exists(temp_audio):
                try:
                    os.remove(temp_audio)
                except OSError:
                    pass
            if temp_wav and os.path.exists(temp_wav):
                try:
                    os.remove(temp_wav)
                except OSError:
                    pass

    def _generate_openai_tts_file(self, text, output_path):
        response = self.openai_client.audio.speech.create(
            model=self.openai_tts_model,
            voice=self.openai_tts_voice,
            input=text,
            instructions=self.openai_tts_instructions,
            response_format="mp3",
            speed=self.openai_tts_speed,
        )
        temp_audio = f"{output_path}.mp3"
        try:
            response.write_to_file(temp_audio)
            subprocess.run(
                [
                    self.ffmpeg_path,
                    "-y",
                    "-i",
                    temp_audio,
                    "-ac",
                    "1",
                    "-ar",
                    "16000",
                    output_path,
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=True,
                timeout=20,
            )
        finally:
            if os.path.exists(temp_audio):
                try:
                    os.remove(temp_audio)
                except OSError:
                    pass

    def _speak_with_flite(self, text):
        prompt = self._prepare_tts_text(text)
        temp_wav = None
        try:
            with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as tmp:
                temp_wav = tmp.name
            subprocess.run(
                [
                    self.ffmpeg_path,
                    "-y",
                    "-f",
                    "lavfi",
                    "-i",
                    f"flite=text='{prompt}':voice={self.flite_voice}",
                    "-filter:a",
                    f"atempo={self.flite_tempo}",
                    "-ar",
                    "16000",
                    "-ac",
                    "1",
                    temp_wav,
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=True,
                timeout=15,
            )
            winsound.PlaySound(temp_wav, winsound.SND_FILENAME)
        finally:
            if temp_wav and os.path.exists(temp_wav):
                try:
                    os.remove(temp_wav)
                except OSError:
                    pass

    def _generate_flite_tts_file(self, text, output_path):
        prompt = self._prepare_tts_text(text)
        subprocess.run(
            [
                self.ffmpeg_path,
                "-y",
                "-f",
                "lavfi",
                "-i",
                f"flite=text='{prompt}':voice={self.flite_voice}",
                "-filter:a",
                f"atempo={self.flite_tempo}",
                "-ar",
                "16000",
                "-ac",
                "1",
                output_path,
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            check=True,
            timeout=15,
        )

    def warmup(self):
        try:
            if self.tts_mode == "openai_tts" and self.openai_client is not None:
                temp_wav = None
                try:
                    with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as tmp:
                        temp_wav = tmp.name
                    response = self.openai_client.audio.speech.create(
                        model=self.openai_tts_model,
                        voice=self.openai_tts_voice,
                        input="hello",
                        instructions=self.openai_tts_instructions,
                        response_format="wav",
                        speed=self.openai_tts_speed,
                    )
                    response.write_to_file(temp_wav)
                    logger.info("OpenAI TTS warmup complete.")
                finally:
                    if temp_wav and os.path.exists(temp_wav):
                        try:
                            os.remove(temp_wav)
                        except OSError:
                            pass
                self._prime_tts_cache()
            elif self.tts_mode == "ffmpeg_flite" and self.ffmpeg_path:
                self._speak_with_flite("hello")
                logger.info("TTS warmup complete.")
                self._prime_tts_cache()
        except Exception as e:
            logger.warning(f"TTS warmup failed: {e}")
            if self.ffmpeg_path:
                self.tts_mode = "ffmpeg_flite"
                logger.info("Falling back to FFmpeg flite for call audio during warmup.")
                try:
                    self._prime_tts_cache()
                except Exception as cache_error:
                    logger.warning(f"TTS cache warmup failed: {cache_error}")

    def _prime_tts_cache(self):
        for prompt in self._cacheable_prompts():
            if prompt in self.tts_cache and os.path.exists(self.tts_cache[prompt]):
                continue
            with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as tmp:
                out = tmp.name
            try:
                if self.tts_mode == "openai_tts" and self.openai_client is not None and self.ffmpeg_path:
                    self._generate_openai_tts_file(prompt, out)
                elif self.tts_mode == "ffmpeg_flite" and self.ffmpeg_path:
                    self._generate_flite_tts_file(prompt, out)
                else:
                    if os.path.exists(out):
                        os.remove(out)
                    continue
                self.tts_cache[prompt] = out
            except Exception:
                if os.path.exists(out):
                    try:
                        os.remove(out)
                    except OSError:
                        pass

    def ensure_playback_routing(self):
        """
        TTS uses the current Windows playback device, so we try to switch it to VB-Cable.
        """
        ps_script = r"""
        try {
            if (!(Get-Module -ListAvailable -Name AudioDeviceCmdlets)) {
                Write-Output "MODULE_MISSING"
                exit 0
            }

            Import-Module AudioDeviceCmdlets
            $currentDevice = Get-AudioDevice -Playback
            if ($null -eq $currentDevice) {
                Write-Output "NO_PLAYBACK_DEVICE"
                exit 0
            }

            if ($currentDevice.Name -eq "%PLAYBACK_DEVICE%") {
                Write-Output ("READY:" + $currentDevice.Name)
                exit 0
            }

            $target = Get-AudioDevice -List | Where-Object {
                $_.Type -eq "Playback" -and $_.Name -eq "%PLAYBACK_DEVICE%"
            } | Select-Object -First 1

            if ($null -eq $target) {
                Write-Output ("TARGET_NOT_FOUND:" + $currentDevice.Name)
                exit 0
            }

            Set-AudioDevice -ID $target.ID
            Write-Output ("SET:" + $target.Name)
        } catch {
            Write-Output ("ERROR:" + $_.Exception.Message)
        }
        """

        ps_script = ps_script.replace("%PLAYBACK_DEVICE%", VB_CABLE_PLAYBACK_DEVICE)

        try:
            result = subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
                capture_output=True,
                text=True,
                timeout=15,
            )
            status = (result.stdout or "").strip() or (result.stderr or "").strip()
            if status.startswith("READY:"):
                logger.info(f"Playback routing already correct: {status[6:]}")
                self.playback_routing_ready = True
                return True
            if status.startswith("SET:"):
                logger.info(f"Playback routing updated: {status[4:]}")
                self.playback_routing_ready = True
                return True
            if status == "MODULE_MISSING":
                logger.warning("AudioDeviceCmdlets module not available; cannot auto-set playback routing.")
                self.playback_routing_ready = False
                return False
            if status:
                logger.warning(f"Playback routing check result: {status}")
            self.playback_routing_ready = False
            return False
        except Exception as e:
            logger.warning(f"Playback routing check failed: {e}")
            self.playback_routing_ready = False
            return False

    def _find_input_device(self, device_name):
        mic_list = sr.Microphone.list_microphone_names()
        for i, mic_name in enumerate(mic_list):
            logger.info(f"Mic [{i}]: {mic_name}")
            if device_name and device_name.lower() in mic_name.lower():
                logger.info(f"Using microphone: {mic_name}")
                return i

        logger.warning("Requested input device not found. Using default microphone.")
        return None

    def speak(self, text):
        try:
            logger.info(f"AI Speaking: {text}")
            print(f"AI: {text}")
            cached_path = self.tts_cache.get(text)
            if cached_path and os.path.exists(cached_path):
                winsound.PlaySound(cached_path, winsound.SND_FILENAME)
                time.sleep(0.08)
                return
            if self.tts_mode == "openai_tts":
                try:
                    self._speak_with_openai_tts(text)
                except Exception as openai_tts_error:
                    logger.warning(f"OpenAI TTS failed, falling back to FFmpeg flite: {openai_tts_error}")
                    if self.ffmpeg_path:
                        self.tts_mode = "ffmpeg_flite"
                        self._speak_with_flite(text)
                    else:
                        raise
            elif self.tts_mode == "ffmpeg_flite" and self.ffmpeg_path:
                self._speak_with_flite(text)
            elif self.tts_mode == "sapi" and self.sapi_voice is not None:
                self.sapi_voice.Speak(text)
            elif self.tts_mode == "pyttsx3" and self.tts_engine is not None:
                self.tts_engine.say(text)
                self.tts_engine.runAndWait()
            time.sleep(0.08)

        except Exception as e:
            logger.error(f"TTS error: {e}")
            print(f"TTS Error: {e}")
            try:
                logger.info("Reinitializing TTS after failure.")
                self.tts_engine = None
                self.sapi_voice = None
                self._init_tts()
                if self.tts_mode == "openai_tts":
                    try:
                        self._speak_with_openai_tts(text)
                    except Exception:
                        if self.ffmpeg_path:
                            logger.info("Falling back to FFmpeg flite after OpenAI TTS failure.")
                            self.tts_mode = "ffmpeg_flite"
                            self._speak_with_flite(text)
                        else:
                            raise
                elif self.tts_mode == "ffmpeg_flite" and self.ffmpeg_path:
                    self._speak_with_flite(text)
                elif self.tts_mode == "sapi" and self.sapi_voice is not None:
                    self.sapi_voice.Speak(text)
                elif self.tts_engine is not None:
                    self.tts_engine.say(text)
                    self.tts_engine.runAndWait()
            except Exception as retry_error:
                logger.error(f"TTS retry error: {retry_error}")
                print(f"TTS Retry Error: {retry_error}")

    def _listen_with_windows_speech(self, timeout=8, phrase_time_limit=12, context="free"):
        grammar_setup = """
        $grammar = New-Object System.Speech.Recognition.DictationGrammar
        $engine.LoadGrammar($grammar)
        """
        if context == "quality":
            grammar_setup = """
        $choices = New-Object System.Speech.Recognition.Choices
        foreach ($item in @("1","2","3","4","5","one","two","three","four","five","poor","bad","average","good","excellent")) { [void]$choices.Add($item) }
        $builder = New-Object System.Speech.Recognition.GrammarBuilder
        [void]$builder.Append($choices)
        $grammar = New-Object System.Speech.Recognition.Grammar($builder)
        $engine.LoadGrammar($grammar)
        """
        elif context == "confirm":
            grammar_setup = """
        $choices = New-Object System.Speech.Recognition.Choices
        foreach ($item in @("yes","yeah","correct","right","confirmed","true","no","wrong","incorrect","negative")) { [void]$choices.Add($item) }
        $builder = New-Object System.Speech.Recognition.GrammarBuilder
        [void]$builder.Append($choices)
        $grammar = New-Object System.Speech.Recognition.Grammar($builder)
        $engine.LoadGrammar($grammar)
        """
        elif context == "status":
            grammar_setup = """
        $choices = New-Object System.Speech.Recognition.Choices
        foreach ($item in @("good","very good","fine","great","okay","ok","doing well","not good","bad","busy","tired","sad","sick")) { [void]$choices.Add($item) }
        $builder = New-Object System.Speech.Recognition.GrammarBuilder
        [void]$builder.Append($choices)
        $grammar = New-Object System.Speech.Recognition.Grammar($builder)
        $engine.LoadGrammar($grammar)
        """
        elif context == "cli":
            grammar_setup = """
        $choices = New-Object System.Speech.Recognition.Choices
        foreach ($item in @("zero","one","two","three","four","five","six","seven","eight","nine","oh","0","1","2","3","4","5","6","7","8","9")) { [void]$choices.Add($item) }
        $builder = New-Object System.Speech.Recognition.GrammarBuilder
        [void]$builder.Append($choices, 8, 15)
        $grammar = New-Object System.Speech.Recognition.Grammar($builder)
        $engine.LoadGrammar($grammar)
        """

        ps_script = f"""
Add-Type -AssemblyName System.Speech
$engine = New-Object System.Speech.Recognition.SpeechRecognitionEngine ([System.Globalization.CultureInfo]::GetCultureInfo('en-US'))
$engine.SetInputToDefaultAudioDevice()
$engine.InitialSilenceTimeout = [TimeSpan]::FromSeconds({timeout})
$engine.BabbleTimeout = [TimeSpan]::FromSeconds({timeout})
$engine.EndSilenceTimeout = [TimeSpan]::FromMilliseconds(700)
$engine.EndSilenceTimeoutAmbiguous = [TimeSpan]::FromMilliseconds(900)
{grammar_setup}
$result = $engine.Recognize([TimeSpan]::FromSeconds({max(timeout, phrase_time_limit) + 1}))
if ($result -and $result.Text) {{
    Write-Output $result.Text
}}
"""

        try:
            result = subprocess.run(
                ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
                capture_output=True,
                text=True,
                timeout=max(timeout, phrase_time_limit) + 6,
            )
            text = (result.stdout or "").strip()
            if text:
                text = self._normalize_recognized_text(text, context=context)
                logger.info(f"Caller said: {text}")
                print(f"Caller: {text}")
                return text
            return None
        except subprocess.TimeoutExpired:
            logger.warning(f"Windows speech recognition timed out for context={context}.")
            return None
        except Exception as e:
            logger.error(f"Windows speech recognition error: {e}")
            return None

    def listen(self, timeout=8, phrase_time_limit=12, context="free", expected_length=None):
        try:
            logger.info("Listening for caller response...")
            print("Listening for caller response...")
            time.sleep(0.15)

            clip_path = self._capture_response_audio(timeout=timeout, phrase_time_limit=phrase_time_limit)
            if clip_path:
                processed_path = self._preprocess_audio_for_transcription(clip_path)
                try:
                    text = self._transcribe_with_openai(
                        processed_path,
                        context=context,
                        expected_length=expected_length,
                    )
                    if text:
                        return text
                finally:
                    if processed_path != clip_path and processed_path and os.path.exists(processed_path):
                        try:
                            os.remove(processed_path)
                        except OSError:
                            pass
                    try:
                        os.remove(clip_path)
                    except OSError:
                        pass

                if self.openai_client is not None and context in {"otp", "cli"}:
                    return None

            text = self._listen_with_windows_speech(
                timeout=timeout,
                phrase_time_limit=phrase_time_limit,
                context=context,
            )
            if text:
                return text

            if context in {"quality", "confirm"}:
                text = self._listen_with_windows_speech(
                    timeout=min(timeout, 4),
                    phrase_time_limit=min(phrase_time_limit, 3),
                    context="free",
                )
                if text:
                    text = self._normalize_recognized_text(text, context=context)
                    logger.info(f"Caller said (fallback): {text}")
                    print(f"Caller: {text}")
                    return text

            with sr.Microphone(device_index=self.mic_index) as source:
                self.recognizer.adjust_for_ambient_noise(source, duration=0.3)
                audio = self.recognizer.listen(
                    source,
                    timeout=timeout,
                    phrase_time_limit=min(
                        phrase_time_limit if phrase_time_limit is not None else MAX_DYNAMIC_RESPONSE_SECONDS,
                        MAX_DYNAMIC_RESPONSE_SECONDS,
                    ),
                )

            text = self.recognizer.recognize_google(audio, language="en-US")
            text = self._normalize_recognized_text(text, context=context)
            logger.info(f"Caller said: {text}")
            print(f"Caller: {text}")
            return text
        except sr.WaitTimeoutError:
            logger.warning("Listening timed out while waiting for caller response.")
            return None
        except sr.UnknownValueError:
            logger.warning("Speech was detected but could not be recognized.")
            return None
        except sr.RequestError as e:
            logger.error(f"Speech recognition request failed: {e}")
            return None
        except Exception as e:
            logger.error(f"Listen error: {e}")
            return None

    def ask_and_listen(self, prompt, timeout=8, phrase_time_limit=12, retries=2, listen_timeout=None, context="free", expected_length=None):
        self.speak(prompt)
        for attempt in range(retries + 1):
            response = self.listen(
                timeout=listen_timeout if listen_timeout is not None else timeout,
                phrase_time_limit=phrase_time_limit,
                context=context,
                expected_length=expected_length,
            )
            if response:
                return response
            if attempt < retries:
                self.speak("I am sorry, I did not catch that. Please say it again.")
        return None

    def acknowledge_status(self, user_input):
        text = (user_input or "").lower()
        if any(word in text for word in ["good", "great", "fine", "well", "excellent", "okay", "ok"]):
            return "That is lovely to hear."
        if any(word in text for word in ["bad", "not good", "tired", "busy", "sad", "sick"]):
            return "Thank you for sharing that. I hope your day gets easier."
        return "Thank you for sharing that with me."

    def acknowledge_quality(self, quality_label):
        mapping = {
            "Excellent": "Excellent, thank you so much.",
            "Good": "Good, thank you so much.",
            "Average": "Thank you, I noted average quality.",
            "Bad": "Thank you, I noted bad quality.",
            "Poor": "Thank you, I noted poor quality.",
        }
        return mapping.get(quality_label, "Thank you, I noted that.")

    def generate_otp(self):
        return "".join(str(random.randint(0, 9)) for _ in range(4))

    def format_digits_for_speech(self, digits):
        clean = re.sub(r"\D", "", str(digits or ""))
        return " ".join(list(clean)) if clean else ""

    def extract_exact_digits(self, text, expected_length):
        source = str(text or "").lower()

        raw_digit_groups = re.findall(r"\d+", source)
        for group in raw_digit_groups:
            if len(group) == expected_length:
                return group

        parts = re.findall(r"[A-Za-z0-9]+", source)
        mapped_parts = []
        word_to_digit = {
            "zero": "0", "oh": "0", "o": "0",
            "one": "1", "won": "1",
            "two": "2", "to": "2", "too": "2",
            "three": "3", "tree": "3", "free": "3",
            "four": "4", "for": "4",
            "five": "5",
            "six": "6",
            "seven": "7",
            "eight": "8", "ate": "8",
            "nine": "9",
        }
        for part in parts:
            mapped_parts.append(word_to_digit.get(part, ""))

        digits = "".join(mapped_parts)
        if len(digits) == expected_length:
            return digits
        return None

    def confirm_yes_no(self, prompt, timeout, phrase_time_limit):
        response = self.ask_and_listen(
            prompt,
            timeout=timeout,
            phrase_time_limit=phrase_time_limit,
            retries=1,
            context="confirm",
        )
        if not response:
            return False, None
        normalized = response.lower()
        if any(word in normalized for word in ["yes", "yeah", "correct", "right", "confirmed", "true"]):
            return True, response
        if any(word in normalized for word in ["no", "wrong", "incorrect", "not", "negative"]):
            return False, response
        return False, response

    def confirm_recording_consent(self):
        confirmed, response = self.confirm_yes_no(
            "This call is being recorded for quality purposes. Do I have your permission to continue?",
            timeout=5,
            phrase_time_limit=4,
        )
        return confirmed, response

    def confirm_user_availability(self):
        confirmed, response = self.confirm_yes_no(
            "Are you free and available for a quick voice quality test?",
            timeout=5,
            phrase_time_limit=4,
        )
        return confirmed, response

    def confirm_otp_twice(self, otp_digits):
        spoken_otp = self.format_digits_for_speech(otp_digits)
        confirmations = []

        for prompt in [
            f"Your one time password is {spoken_otp}. Please repeat the one time password back to me.",
            f"Thank you. Please repeat the one time password once more.",
        ]:
            response = self.ask_and_listen(
                prompt,
                timeout=6,
                phrase_time_limit=5,
                retries=1,
                listen_timeout=7,
                context="otp",
            )
            repeated_digits = self.extract_exact_digits(response, 4) if response else None
            confirmed = bool(repeated_digits and repeated_digits == otp_digits)
            confirmations.append(
                {
                    "confirmed": confirmed,
                    "response": response,
                    "repeated_digits": repeated_digits or "Not captured",
                }
            )
            if not confirmed:
                return False, confirmations

        return True, confirmations

    def extract_audio_quality(self, text):
        if not text:
            return None
        text_lower = text.lower()

        for num in ["5", "five"]:
            if re.search(rf"\b{num}\b", text_lower):
                return "Excellent"
        for num in ["4", "four"]:
            if re.search(rf"\b{num}\b", text_lower):
                return "Good"
        for num in ["3", "three"]:
            if re.search(rf"\b{num}\b", text_lower):
                return "Average"
        for num in ["2", "two"]:
            if re.search(rf"\b{num}\b", text_lower):
                return "Bad"
        for num in ["1", "one"]:
            if re.search(rf"\b{num}\b", text_lower):
                return "Poor"

        if any(word in text_lower for word in ["excellent", "perfect", "great", "amazing"]):
            return "Excellent"
        if any(word in text_lower for word in ["good", "nice", "clear", "fine"]):
            return "Good"
        if any(word in text_lower for word in ["okay", "ok", "average", "decent"]):
            return "Average"
        if any(word in text_lower for word in ["bad", "weak", "unclear"]):
            return "Bad"
        if any(word in text_lower for word in ["poor", "bad", "terrible", "awful"]):
            return "Poor"

        return None

    def _normalize_digits_from_speech(self, text):
        word_to_digit = {
            "zero": "0", "oh": "0", "o": "0",
            "one": "1", "won": "1",
            "two": "2", "to": "2", "too": "2",
            "three": "3", "tree": "3", "free": "3",
            "four": "4", "for": "4",
            "five": "5",
            "six": "6",
            "seven": "7",
            "eight": "8", "ate": "8",
            "nine": "9",
        }

        parts = re.findall(r"[A-Za-z0-9]+", text.lower())
        converted = []
        for p in parts:
            converted.append(word_to_digit.get(p, p))

        return re.sub(r"\D", "", "".join(converted))

    def extract_cli_number(self, text, expected_length=None):
        if not text:
            return None

        spoken_digits = self._normalize_digits_from_speech(text)
        if expected_length:
            if len(spoken_digits) == expected_length:
                return spoken_digits
            return self._extract_plausible_cli_digits(spoken_digits)
        plausible_digits = self._extract_plausible_cli_digits(spoken_digits)
        if plausible_digits:
            return plausible_digits

        cleaned = text.replace(" ", "").replace("-", "")
        patterns = [
            r"\+?966\d{9}",
            r"\d{12}",
            r"\d{10}",
            r"\d{9}",
        ]

        for pattern in patterns:
            matches = re.findall(pattern, cleaned)
            if matches:
                candidate = re.sub(r"[^\d]", "", matches[0])
                if expected_length and len(candidate) != expected_length:
                    candidate = self._extract_plausible_cli_digits(candidate)
                else:
                    candidate = self._extract_plausible_cli_digits(candidate) or candidate
                if candidate:
                    return candidate

        return None

    def _extract_plausible_cli_digits(self, digits):
        clean = re.sub(r"\D", "", str(digits or ""))
        if not clean:
            return None

        groups = re.findall(r"\d{9,10}", clean)
        if groups:
            return max(groups, key=len)

        if len(clean) > 10:
            return clean[-10:]
        if len(clean) == 9:
            return clean
        return None

    def _expected_cli_candidates(self, expected_cli):
        digits = re.sub(r"\D", "", str(expected_cli or ""))
        candidates = []
        if len(digits) >= 10:
            candidates.append(digits[-10:])
        if len(digits) >= 9:
            candidates.append(digits[-9:])
        if 9 <= len(digits) <= 10:
            candidates.append(digits)
        # preserve order while removing duplicates
        return list(dict.fromkeys([c for c in candidates if c]))

    def run_conversation(self, expected_cli, max_duration=120):
        start_time = time.time()
        routing_ok = self.ensure_playback_routing()
        if not routing_ok:
            logger.warning("Continuing without confirmed VB-Cable playback routing.")

        # Keep the greeting nearly immediate while still allowing a brief media-path settle.
        time.sleep(0.4)

        conversation_turns = 0
        recording_consent = "Denied"
        user_available = "No"

        self.speak("Hello, this is Monica from Mobily Voice Quality Testing. It is so nice to speak with you today.")

        consent_ok, consent_response = self.confirm_recording_consent()
        if consent_response:
            conversation_turns += 1
            self.conversation_history.append({"role": "user", "content": consent_response})

        if not consent_ok:
            self.speak("Thank you. Since recording permission was not granted, I will end the call now. Goodbye.")
            return {
                "audio_quality": "Not captured",
                "caller_status": "Recording denied",
                "cli_reported": "Not captured",
                "cli_expected": expected_cli,
                "cli_match": "Mismatch",
                "conversation_turns": conversation_turns,
                "duration_seconds": int(time.time() - start_time),
                "recording_consent": "Denied",
                "otp_value": "Not generated",
                "otp_confirmed": False,
                "end_reason": "recording_denied",
            }

        recording_consent = "Confirmed"
        available_ok, available_response = self.confirm_user_availability()
        if available_response:
            conversation_turns += 1
            self.conversation_history.append({"role": "user", "content": available_response})

        if not available_ok:
            self.speak("Thank you. Since this is not a good time, I will end the call now. Goodbye.")
            return {
                "audio_quality": "Not captured",
                "caller_status": recording_consent,
                "cli_reported": "Not captured",
                "cli_expected": expected_cli,
                "cli_match": "Mismatch",
                "conversation_turns": conversation_turns,
                "duration_seconds": int(time.time() - start_time),
                "recording_consent": recording_consent,
                "user_available": "No",
                "end_reason": "user_not_available",
            }

        user_available = "Yes"
        self.speak("Thank you. Let us continue with the voice quality test.")

        expected_cli_candidates = self._expected_cli_candidates(expected_cli)
        expected_cli_length = len(expected_cli_candidates[0]) if expected_cli_candidates else None

        spoken_cli = None
        spoken_cli_response = None
        for _ in range(2):
            spoken_cli_response = self.ask_and_listen(
                "Please confirm the caller I D number you see on your screen, digit by digit.",
                timeout=5,
                phrase_time_limit=5,
                retries=1,
                listen_timeout=5,
                context="cli",
                expected_length=expected_cli_length,
            )
            spoken_cli = (
                self.extract_cli_number(
                    spoken_cli_response,
                    expected_length=expected_cli_length,
                )
                if spoken_cli_response
                else None
            )
            if spoken_cli_response:
                conversation_turns += 1
                self.conversation_history.append({"role": "user", "content": spoken_cli_response})
            if not spoken_cli:
                self.speak("I could not capture the full caller I D. Could you please repeat the complete number slowly, digit by digit?")
                continue

            confirmed, confirmation_text = self.confirm_yes_no(
                f"I heard caller I D {self.format_digits_for_speech(spoken_cli)}. Is that correct?",
                timeout=8,
                phrase_time_limit=6,
            )
            if confirmation_text:
                conversation_turns += 1
                self.conversation_history.append({"role": "user", "content": confirmation_text})
            if confirmed:
                self.speak("Perfect, thank you for confirming the caller I D.")
                break
            self.speak("Could you please repeat the caller I D once more?")
            spoken_cli = None

        if not spoken_cli:
            self.speak("Thank you. I was not able to confirm the caller I D, but we can continue.")

        audio_quality = None
        audio_quality_response = None
        for _ in range(2):
            audio_quality_response = self.ask_and_listen(
                "Please share your voice quality feedback. You may say excellent, good, average, bad, or poor.",
                timeout=4,
                phrase_time_limit=4,
                retries=1,
                listen_timeout=6,
                context="quality",
            )
            audio_quality = self.extract_audio_quality(audio_quality_response) if audio_quality_response else None
            if audio_quality_response:
                conversation_turns += 1
                self.conversation_history.append({"role": "user", "content": audio_quality_response})
            if not audio_quality:
                self.speak("I could not understand the feedback. Please say excellent, good, average, bad, or poor.")
                continue

            self.speak(self.acknowledge_quality(audio_quality))
            confirmed, confirmation_text = self.confirm_yes_no(
                f"I understood the call quality as {audio_quality}. Is that correct?",
                timeout=4,
                phrase_time_limit=3,
            )
            if confirmation_text:
                conversation_turns += 1
                self.conversation_history.append({"role": "user", "content": confirmation_text})
            if confirmed:
                break
            self.speak("Thank you. Please share the voice quality once more.")
            audio_quality = None

        self.speak("Thank you very much for your time. Goodbye.")

        cli_expected_digits = re.sub(r"\D", "", str(expected_cli or ""))
        cli_reported_digits = re.sub(r"\D", "", str(spoken_cli or ""))
        cli_match = "Not confirmed"
        if cli_reported_digits:
            cli_match = "Match" if cli_reported_digits in expected_cli_candidates else "Mismatch"

        return {
            "audio_quality": audio_quality or "Not captured",
            "caller_status": recording_consent,
            "cli_reported": spoken_cli or "Not captured",
            "cli_expected": expected_cli,
            "cli_match": cli_match,
            "conversation_turns": conversation_turns,
            "duration_seconds": int(time.time() - start_time),
            "recording_consent": recording_consent,
            "user_available": user_available,
            "end_reason": "completed",
        }

    def cleanup(self):
        try:
            if self.tts_engine:
                self.tts_engine.stop()
        except Exception:
            pass
