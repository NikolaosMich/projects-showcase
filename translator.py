import pyaudio
import wave
import speech_recognition as sr
import threading
import queue
import io
import translators as ts

# Initialize PyAudio
p = pyaudio.PyAudio()

# ---------------------------------- Input Device Selection ----------------------------------
# This section identifies and selects the desired input device for audio capture.
# In this example, the "Stereo Mix" device is selected. Adjust the device name or setup based on your system configuration.

stereo_mix_device_index = None

# Iterate over all available audio devices to find "Stereo Mix".
for i in range(p.get_device_count()):
    info = p.get_device_info_by_index(i)
    print(info)  # Optional: prints device information to help identify the input device.
    if "Stereo Mix" in info['name']:
        stereo_mix_device_index = i
        break

if stereo_mix_device_index is not None:
    print(f"'Stereo Mix' found at device index {stereo_mix_device_index}")
else:
    print("No 'Stereo Mix' device found. Please ensure the device is enabled or use an alternative input.")

# ---------------------------------- Audio Configuration ----------------------------------
# Configuration for audio recording
FORMAT = pyaudio.paInt16
CHANNELS = 1
RATE = 14000
CHUNK = 2048
BUFFER_SECONDS = 20
DEVICE_INDEX = stereo_mix_device_index

recognizer = sr.Recognizer()
audio_queue = queue.Queue()
recording = True


# ---------------------------------- Audio Recording Function ----------------------------------
def record_audio():
    """Captures audio from the selected input device and stores it in a queue."""
    try:
        # Open the input stream with the selected device and settings
        stream = p.open(format=FORMAT,
                        channels=CHANNELS,
                        rate=RATE,
                        input=True,
                        input_device_index=DEVICE_INDEX,
                        frames_per_buffer=CHUNK)

        print("Recording started...")
        while recording:
            data = stream.read(CHUNK)
            audio_queue.put(data)

    except Exception as e:
        print(f"Error during recording: {e}")
    finally:
        stream.stop_stream()
        stream.close()


# ---------------------------------- Audio Processing Function ----------------------------------
def process_audio():
    """Processes the buffered audio, recognizes speech, and translates it."""
    buffer = []
    while recording or not audio_queue.empty():
        if not audio_queue.empty():
            try:
                # Collect audio chunks until BUFFER_SECONDS worth of audio is gathered
                while len(buffer) < int(RATE / CHUNK * BUFFER_SECONDS):
                    buffer.append(audio_queue.get())

                audio_data = b''.join(buffer)
                audio_file = io.BytesIO()  # Create an in-memory WAV file

                # Write the audio data to the in-memory WAV file
                with wave.open(audio_file, 'wb') as wf:
                    wf.setnchannels(CHANNELS)
                    wf.setsampwidth(p.get_sample_size(FORMAT))
                    wf.setframerate(RATE)
                    wf.writeframes(audio_data)

                # Keep a small amount of data from the buffer to overlap with the next round of processing
                buffer = buffer[-4:]

                # Process the in-memory audio file for speech recognition
                audio_file.seek(0)
                with sr.AudioFile(audio_file) as source:
                    audio = recognizer.record(source)
                    try:
                        text = recognizer.recognize_whisper(audio, model="medium", language="german")
                        print(f"Recognized (German): {text}")
                    except sr.UnknownValueError:
                        print("Speech recognition could not understand the audio.")
                    except sr.RequestError as e:
                        print(f"Error with speech recognition service: {e}")

                translated_text = ts.translate_text(text, from_language="de")
                print(f"Translated (to English): {translated_text}")

            except Exception as e:
                print(f"Error during audio processing: {e}")

# ---------------------------------- Multithreading for Recording and Processing ----------------------------------
record_thread = threading.Thread(target=record_audio)
process_thread = threading.Thread(target=process_audio)

record_thread.start()
process_thread.start()

# Wait for both threads to complete (join)
record_thread.join()
process_thread.join()

# Terminate PyAudio when done
p.terminate()