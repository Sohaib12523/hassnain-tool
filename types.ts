


export type AspectRatio = "1:1" | "16:9" | "9:16" | "4:3" | "3:4";
export type Mode = 'bulk' | 'script-to-scene' | 'character' | 'story-weaver' | 'voice-generator' | 'ai-script-tuner' | 'bulk-script-writer';
export type StoryWeaverMode = 'idea' | 'script' | 'voice';
export type VoiceGeneratorMode = 'voice' | 'sfx';
export type ApiProvider = 'gemini' | 'together' | 'leonardo' | 'ideogram' | 'runway';
export type BulkScriptApiProvider = 'gemini' | 'openai' | 'groq' | 'deepseek';
export type VoiceProvider = 'elevenlabs' | 'wellsaid';


export type CaptionAnimation = 'none' | 'fade-in' | 'slide-up' | 'pop-in';
export type SceneTransition = 'none' | 'fade' | 'slide-left';
export type PreviewState = 'idle' | 'loading' | 'playing';

export interface ImageJob {
    id: string;
    prompt: string;
    imageUrl?: string;
    error?: string;
}

export interface BulkScriptJob {
    id: string;
    title: string;
    guides?: string;
    outline?: string;
    wordCount: number;
    script?: string;
    status: 'pending' | 'analyzing' | 'review' | 'generating' | 'complete' | 'error';
    error?: string;
}

export interface CharacterDetails {
    description: string;
    refImage: {
        url: string; // data URL
        base64: string;
    } | null;
}

export interface GenerationResult {
    prompt: string;
    url?: string;
    error?: string;
}

export interface GenerationParams {
    prompts: string[];
    ratio: AspectRatio;
    steps?: number; // Optional: only used by Together AI service
    onProgress: (index: number, result: GenerationResult) => void;
    apiKey?: string | null;
}

export interface VideoConfig {
    voice: string;
    captionsEnabled: boolean;
    backgroundEnabled: boolean;
    strokeEnabled: boolean;
    shadowEnabled: boolean;
    captionColor: string;
    captionStrokeColor: string;
    captionStrokeWidth: number;
    captionShadowColor: string;
    captionShadowBlur: number;
    captionBackgroundColor: string;
    captionBackgroundOpacity: number;
    captionWordsPerLine: number;
    captionAnimation: CaptionAnimation;
    captionFont: string;
    sceneTransition: SceneTransition;
}

export interface ElevenLabsVoice {
  voice_id: string;
  name: string;
  category: string;
  labels: Record<string, string>;
}

export interface Voice {
    id: string;
    name: string;
    provider: VoiceProvider;
    group?: string; // category for ElevenLabs, style for WellSaid
}

export interface VoiceScript {
    id: string;
    text: string;
    voiceId: string;
    provider: VoiceProvider;
    speed: number; // ElevenLabs specific
    isGenerating?: boolean;
    audioUrl?: string;
    error?: string;
    isModified?: boolean;
    justUpdated?: boolean;
}

export interface SoundEffectJob {
    id: string;
    text: string;
    isGenerating?: boolean;
    audioUrl?: string;
    error?: string;
}

export interface AllApiKeys {
  image: Partial<Record<ApiProvider, string>>;
  voice: Partial<Record<VoiceProvider, string>>;
  script: Partial<Record<Exclude<BulkScriptApiProvider, 'gemini'>, string>>;
}