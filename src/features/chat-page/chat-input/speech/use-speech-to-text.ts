import { showError } from "@/features/globals/global-message-store";
import {
  AudioConfig,
  AutoDetectSourceLanguageConfig,
  SpeechConfig,
  SpeechRecognizer,
} from "microsoft-cognitiveservices-speech-sdk";
import { proxy, useSnapshot } from "valtio";
import { chatStore } from "../../chat-store";
import { GetSpeechToken } from "./speech-service";

let speechRecognizer: SpeechRecognizer | undefined = undefined;
let confirmedText = "";
let isStopping = false;
let isStarting = false;
let currentSessionId = 0;

class SpeechToText {
  public isMicrophoneUsed: boolean = false;
  public isMicrophoneReady: boolean = false;

  public async startRecognition() {
    // トークン取得中・停止中・認識中はすべてガード
    if (isStarting || isStopping || speechRecognizer) return;
    isStarting = true;

    const token = await GetSpeechToken();

    if (token.error) {
      showError(token.errorMessage);
      isStarting = false;
      return;
    }

    // セッションID を採番（各コールバックが古いセッションを無視するために使用）
    const mySession = ++currentSessionId;

    this.isMicrophoneReady = true;
    this.isMicrophoneUsed = true;
    confirmedText = chatStore.input;

    const speechConfig = SpeechConfig.fromAuthorizationToken(
      token.token,
      token.region
    );

    const audioConfig = AudioConfig.fromDefaultMicrophoneInput();

    const autoDetectSourceLanguageConfig =
      AutoDetectSourceLanguageConfig.fromLanguages([

        "ja-JP",
      ]);

    const recognizer = SpeechRecognizer.FromConfig(
      speechConfig,
      autoDetectSourceLanguageConfig,
      audioConfig
    );

    speechRecognizer = recognizer;
    isStarting = false;

    // 息継ぎ区切りの「。」を「、」に変換（会話終了時に stopRecognition で句点に戻す）
    const toComma = (t: string) => t.replace(/。$/, "、");

    recognizer.recognizing = (s, e) => {
      if (currentSessionId !== mySession || !e.result.text) return;
      const interim = toComma(e.result.text);
      chatStore.updateInput(
        confirmedText ? confirmedText + "\n" + interim : interim
      );
    };

    recognizer.recognized = (s, e) => {
      if (currentSessionId !== mySession) return;
      if (e.result.text) {
        confirmedText += (confirmedText ? "\n" : "") + toComma(e.result.text);
        chatStore.updateInput(confirmedText);
      }
    };

    recognizer.canceled = (s, e) => {
      // 古いセッションのキャンセルは close のみ（不要なエラー表示を防ぐ）
      if (currentSessionId !== mySession) {
        recognizer.close();
        return;
      }
      showError(e.errorDetails);
      this.isMicrophoneReady = false;
      speechRecognizer = undefined;
      confirmedText = "";
      isStopping = false;
      recognizer.close();
    };

    recognizer.startContinuousRecognitionAsync();
  }

  public stopRecognition() {
    if (!speechRecognizer || isStopping) return;
    isStopping = true;
    this.isMicrophoneReady = false;
    const r = speechRecognizer;
    const stoppingSession = currentSessionId;
    r.stopContinuousRecognitionAsync(() => {
      // 停止完了までに新セッションが始まっていたら confirmedText には触らない
      if (currentSessionId !== stoppingSession) {
        isStopping = false;
        r.close();
        return;
      }
      // 会話終了時: 末尾の読点を句点に変換
      const finalText = confirmedText.replace(/、$/, "。");
      if (finalText !== confirmedText) {
        chatStore.updateInput(finalText);
      }
      speechRecognizer = undefined;
      confirmedText = "";
      isStopping = false;
      r.close();
    });
  }

  public userDidUseMicrophone() {
    return this.isMicrophoneUsed;
  }

  public resetMicrophoneUsed() {
    this.isMicrophoneUsed = false;
  }
}
export const speechToTextStore = proxy(new SpeechToText());

export const useSpeechToText = () => {
  return useSnapshot(speechToTextStore);
};
