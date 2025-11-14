import { showError } from "@/features/globals/global-message-store";
import {
  AudioConfig,
  ResultReason,
  SpeakerAudioDestination,
  SpeechConfig,
  SpeechSynthesizer,
} from "microsoft-cognitiveservices-speech-sdk";
import { proxy, useSnapshot } from "valtio";
import { GetSpeechToken } from "./speech-service";
import { speechToTextStore } from "./use-speech-to-text";

let player: SpeakerAudioDestination | undefined = undefined;

class TextToSpeech {
  public isPlaying: boolean = false;

  public stopPlaying() {
    this.isPlaying = false;
    if (player) {
      player.pause();
    }
  }

  // 言語判定のための関数
  private detectLanguage(text: string): string {
    const japaneseRegex = /[\u3000-\u30FF\u4E00-\u9FAF\uFF66-\uFF9F]/; // 日本語を判定する簡易的な正規表現
    return japaneseRegex.test(text) ? "ja-JP" : "en-US"; // 日本語なら"ja-JP"、英語なら"en-US"
  }

  public async textToSpeech(textToSpeak: string) {
    if (this.isPlaying) {
      this.stopPlaying();
    }

    // URLとハッシュタグを削除する正規表現
    const filteredText = textToSpeak
      .replace(/https?:\/\/[^\s]+/g, '') // URLを削除
      .replace(/#[^\s]+/g, ''); // ハッシュタグを削除

    const tokenObj = await GetSpeechToken();

    if (tokenObj.error) {
      showError(tokenObj.errorMessage);
      return;
    }

    // 言語を判定して、適切な音声を選択
    const language = this.detectLanguage(filteredText);
    let voiceName = "ja-JP-NanamiNeural"; // デフォルトは日本語

    if (language === "en-US") {
      voiceName = "en-US-JennyNeural"; // 英語の場合は英語の音声を指定
    }

    const speechConfig = SpeechConfig.fromAuthorizationToken(
      tokenObj.token,
      tokenObj.region
    );
    speechConfig.speechSynthesisVoiceName = voiceName;

    player = new SpeakerAudioDestination();

    const audioConfig = AudioConfig.fromSpeakerOutput(player);
    const synthesizer = new SpeechSynthesizer(speechConfig, audioConfig);

    player.onAudioEnd = () => {
      this.isPlaying = false;
    };

    synthesizer.speakTextAsync(
      filteredText, // フィルタリングされたテキストを使用
      (result) => {
        if (result.reason === ResultReason.SynthesizingAudioCompleted) {
          this.isPlaying = true;
        } else {
          showError(result.errorDetails);
          this.isPlaying = false;
        }
        synthesizer.close();
      },
      function (err) {
        console.error("err - " + err);
        synthesizer.close();
      }
    );
  }

  public speak(value: string) {
    if (speechToTextStore.userDidUseMicrophone()) {
      textToSpeechStore.textToSpeech(value);
      speechToTextStore.resetMicrophoneUsed();
    }
  }
}

export const textToSpeechStore = proxy(new TextToSpeech());

export const useTextToSpeech = () => {
  return useSnapshot(textToSpeechStore);
};
