
import React, { useState, useCallback, useRef, useEffect, useMemo } from 'react';
import JSZip from 'jszip';
import mammoth from 'mammoth';
import * as docx from 'docx';
import saveAs from 'file-saver';
import * as multiProviderService from './services/aiService';
import * as geminiService from './services/geminiService';
import * as leonardoService from './services/leonardoService';
import * as runwayService from './services/runwayService';
import * as ideogramService from './services/ideogramService';
import * as elevenLabsService from './services/elevenLabsService';
import * as wellSaidLabsService from './services/wellSaidLabsService';


import { AspectRatioIcon, DownloadIcon, FilmIcon, GeminiIcon, PhotoIcon, RegenerateIcon, ScriptIcon, TogetherIcon, ZipIcon, EditIcon, UserCircleIcon, CheckCircleIcon, ClearIcon, StoryIcon, VideoIcon, SettingsIcon, UndoIcon, RedoIcon, TrashIcon, CloseIcon, LogoIcon, SparkleIcon, LightbulbIcon, DocumentTextIcon, MicrophoneIcon, UploadIcon, SunIcon, RightArrowIcon, PlayIcon, MoonIcon, KeyIcon, WandIcon, CreditIcon, StopIcon, SoundWaveIcon, TuneIcon, ClipboardIcon, BulkScriptIcon, TxtIcon, DocxIcon } from './components/Icons';
import EditModal from './components/ImageModal';
import ApiKeyModal from './components/ApiKeyModal';
import CreditDisplay from './components/CreditDisplay';
import type { ApiProvider, AspectRatio, ImageJob, GenerationResult, Mode, VideoConfig, CaptionAnimation, SceneTransition, StoryWeaverMode, CharacterDetails, Voice, PreviewState, VoiceScript, SoundEffectJob, BulkScriptJob, VoiceGeneratorMode, BulkScriptApiProvider, VoiceProvider, AllApiKeys } from './types';
import Spinner from './components/Spinner';
import VoiceSelector from './components/VoiceSelector';

const ratioMap: Record<AspectRatio, string> = {
    '16:9': 'aspect-[16/9]',
    '9:16': 'aspect-[9/16]',
    '1:1': 'aspect-square',
    '4:3': 'aspect-[4/3]',
    '3:4': 'aspect-[3/4]',
};

type Theme = 'light' | 'dark';

const DEFAULT_IMAGE_STYLES = ['None', 'Cinematic', 'Photographic', 'Anime', 'Digital Art', 'Fantasy', 'Low Poly', 'Pixel Art', 'Comic Book', 'Abstract'];
const ASPECT_RATIOS: AspectRatio[] = ['16:9', '9:16', '1:1', '4:3', '3:4'];
const CAPTION_ANIMATIONS: CaptionAnimation[] = ['none', 'fade-in', 'slide-up', 'pop-in'];
const SCENE_TRANSITIONS: SceneTransition[] = ['none', 'fade', 'slide-left'];
const CAPTION_FONTS = ['Inter', 'Montserrat', 'Oswald', 'Lobster', 'Roboto Condensed'];
const DAILY_CREDITS = 50;

const IMAGE_API_PROVIDERS: { id: ApiProvider; name: string; icon: JSX.Element}[] = [
    { id: 'gemini', name: 'Gemini', icon: <GeminiIcon /> },
    { id: 'together', name: 'TogetherAI', icon: <TogetherIcon /> },
    { id: 'leonardo', name: 'Leonardo.Ai', icon: <div className="h-4 w-4 rounded-full bg-purple-500" /> },
    { id: 'ideogram', name: 'Ideogram', icon: <div className="h-4 w-4 rounded-full bg-red-500" /> },
    { id: 'runway', name: 'RunwayML', icon: <div className="h-4 w-4 rounded-full bg-black" /> },
];

const BULK_SCRIPT_PROVIDERS: { id: BulkScriptApiProvider; name: string; }[] = [
    { id: 'gemini', name: 'Gemini (Default)' },
    { id: 'groq', name: 'Grok AI' },
    { id: 'openai', name: 'OpenAI (GPT)' },
    { id: 'deepseek', name: 'DeepSeek' },
];

const imageServices = {
    gemini: geminiService,
    together: multiProviderService,
    leonardo: leonardoService,
    runway: runwayService,
    ideogram: ideogramService
};


const initialVideoConfig: VideoConfig = {
    voice: '',
    captionsEnabled: true,
    backgroundEnabled: true,
    strokeEnabled: true,
    shadowEnabled: true,
    captionColor: '#FFFFFF',
    captionStrokeColor: '#000000',
    captionStrokeWidth: 2,
    captionShadowColor: '#000000',
    captionShadowBlur: 5,
    captionBackgroundColor: '#000000',
    captionBackgroundOpacity: 0.5,
    captionWordsPerLine: 8,
    captionAnimation: 'fade-in',
    captionFont: 'Inter',
    sceneTransition: 'fade'
};

const TABS = [
    {key: 'story-weaver', label: 'Storyboard', icon: <StoryIcon/>},
    {key: 'bulk', label: 'Bulk Prompts', icon: <PhotoIcon/>},
    {key: 'script-to-scene', label: 'Script to Scene', icon: <FilmIcon/>},
    {key: 'character', label: 'Character Sheet', icon: <UserCircleIcon/>},
    {key: 'bulk-script-writer', label: 'Bulk Scripts', icon: <BulkScriptIcon/>},
    {key: 'voice-generator', label: 'AI Voice Generator', icon: <SoundWaveIcon />},
    {key: 'ai-script-tuner', label: 'AI Script Tuner', icon: <TuneIcon />},
] as const;


// Simple component to render bold markdown
const SimpleMarkdown: React.FC<{ text: string; className?: string }> = ({ text, className }) => {
  const parts = text.split(/(\*\*.*?\*\*)/g);
  return (
    <pre className={`whitespace-pre-wrap font-sans ${className}`}>
      {parts.map((part, index) => {
        if (part.startsWith('**') && part.endsWith('**')) {
          return <strong key={index} className="text-gray-900 dark:text-white">{part.slice(2, -2)}</strong>;
        }
        return part;
      })}
    </pre>
  );
};

// A properly aligned loading indicator for the bulk script generator buttons.
const BulkScriptLoadingIndicator: React.FC = () => (
    <div className="flex flex-col items-center justify-center gap-2">
        <svg className="animate-spin h-8 w-8 text-white" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
            <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4"></circle>
            <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
        </svg>
        <span className="font-bold">Generating...</span>
        <p className="text-sm opacity-80">Your content is being created...</p>
    </div>
);

// Moved from renderBulkScriptResults to prevent re-creation on every render
const StatusBadge: React.FC<{ status: BulkScriptJob['status'] }> = ({ status }) => {
    const config = {
        pending: { text: 'Pending', color: 'bg-gray-400' },
        analyzing: { text: 'Analyzing', color: 'bg-yellow-500 animate-pulse' },
        review: { text: 'Ready to Generate', color: 'bg-blue-500' },
        generating: { text: 'Generating Script', color: 'bg-cyan-500 animate-pulse' },
        complete: { text: 'Complete', color: 'bg-green-500' },
        error: { text: 'Error', color: 'bg-red-500' },
    }[status];

    return (
        <div className="flex items-center gap-2 text-xs font-medium">
            <span className={`h-2.5 w-2.5 rounded-full ${config.color}`}></span>
            <span className="text-gray-600 dark:text-gray-400">{config.text}</span>
        </div>
    );
};


const App: React.FC = () => {
    // Core State
    const [promptsInput, setPromptsInput] = useState<string>('');
    const [scriptInput, setScriptInput] = useState<string>('');
    const [numScenes, setNumScenes] = useState<number>(5);
    const [mode, setMode] = useState<Mode>('story-weaver');
    const [apiProvider, setApiProvider] = useState<ApiProvider>('gemini');
    
    // Generation Config State
    const [ratio, setRatio] = useState<AspectRatio>('16:9');
    const [steps, setSteps] = useState<number>(10);
    const [imageStyle, setImageStyle] = useState<string>('Cinematic');
    const [imageStyles, setImageStyles] = useState<string[]>(DEFAULT_IMAGE_STYLES);
    const [styleReferenceImage, setStyleReferenceImage] = useState<{ url: string; base64: string } | null>(null);
    
    // UI/Loading State
    const [isGenerating, setIsGenerating] = useState<boolean>(false);
    const [isGeneratingScenes, setIsGeneratingScenes] = useState<boolean>(false);
    const [progress, setProgress] = useState(0);
    const [error, setError] = useState<string | null>(null);
    const [isPanelOpen, setIsPanelOpen] = useState(false); // Default to closed for the new landing page
    const [hasStarted, setHasStarted] = useState(false);
    const [theme, setTheme] = useState<Theme>('dark');
    const [isDescribingStyle, setIsDescribingStyle] = useState<boolean>(false);


    // Results & Editing State
    const [storyWeaverResults, setStoryWeaverResults] = useState<ImageJob[]>([]);
    const [bulkResults, setBulkResults] = useState<ImageJob[]>([]);
    const [scriptToSceneResults, setScriptToSceneResults] = useState<ImageJob[]>([]);
    const [editingJob, setEditingJob] = useState<ImageJob | null>(null);

    // File Input Refs
    const promptsFileInputRef = useRef<HTMLInputElement>(null);
    const scriptFileInputRef = useRef<HTMLInputElement>(null);
    const styleReferenceFileInputRef = useRef<HTMLInputElement>(null);
    const audioFileInputRef = useRef<HTMLInputElement>(null);
    const sfxFileInputRef = useRef<HTMLInputElement>(null);
    const scriptTunerFileInputRef = useRef<HTMLInputElement>(null);
    const characterRefImageRefs = useRef<Record<string, HTMLInputElement | null>>({});
    const voiceScriptFileInputRefs = useRef<Record<string, HTMLInputElement | null>>({});
    const bulkScriptTitlesFileInputRef = useRef<HTMLInputElement>(null);
    const bulkScriptGuidesFileInputRef = useRef<HTMLInputElement>(null);
    const bulkScriptOutlineFileInputRef = useRef<HTMLInputElement>(null);


    // Character Consistency State
    const [characterDescription, setCharacterDescription] = useState<string>('');
    const [characterSheet, setCharacterSheet] = useState<ImageJob[]>([]);
    const [activeCharacter, setActiveCharacter] = useState<ImageJob | null>(null);
    const [isGeneratingCharacter, setIsGeneratingCharacter] = useState<boolean>(false);
    const [extractedCharacters, setExtractedCharacters] = useState<string[]>([]);
    const [mainCharacter, setMainCharacter] = useState<string | null>(null);
    const [characterDetails, setCharacterDetails] = useState<Record<string, CharacterDetails>>({});
    const [isAnalyzingScript, setIsAnalyzingScript] = useState(false);
    const [isDescribingImage, setIsDescribingImage] = useState<string | null>(null);

    // Story Weaver State
    const [storyIdea, setStoryIdea] = useState<string>('');
    const [isWeaving, setIsWeaving] = useState<boolean>(false);
    const [weavingStep, setWeavingStep] = useState<string>('');
    const [storyWeaverMode, setStoryWeaverMode] = useState<StoryWeaverMode>('idea');
    const [videoDuration, setVideoDuration] = useState<number>(1); // in minutes
    const [audioBlob, setAudioBlob] = useState<Blob | null>(null);
    const [isTranscribing, setIsTranscribing] = useState<boolean>(false);
    
    // Video Generation State
    const [isGeneratingVideo, setIsGeneratingVideo] = useState<boolean>(false);
    const [videoUrl, setVideoUrl] = useState<string | null>(null);
    const [audioPreview, setAudioPreview] = useState<HTMLAudioElement | null>(null);
    const [previewStatus, setPreviewStatus] = useState<{ id: string | null; state: PreviewState }>({ id: null, state: 'idle' });
    
    // Undo/Redo state for video config
    const [videoConfigState, setVideoConfigState] = useState({
        past: [] as VideoConfig[],
        present: initialVideoConfig,
        future: [] as VideoConfig[]
    });
    const videoConfig = videoConfigState.present; // Main config object to use for rendering

     // API Key & Credit State
    const [isApiKeyModalOpen, setIsApiKeyModalOpen] = useState<boolean>(false);
    const [customImageApiKeys, setCustomImageApiKeys] = useState<Record<string, string>>({});
    const [tempImageApiKey, setTempImageApiKey] = useState('');
    const [credits, setCredits] = useState<number>(DAILY_CREDITS);
    const [creditResetTime, setCreditResetTime] = useState<Date | null>(null);

    // AI Voice Generator State
    const [voiceGeneratorMode, setVoiceGeneratorMode] = useState<VoiceGeneratorMode>('voice');
    const [voiceProvider, setVoiceProvider] = useState<VoiceProvider>('elevenlabs');
    const [elevenLabsApiKey, setElevenLabsApiKey] = useState<string>('');
    const [wellSaidApiKey, setWellSaidApiKey] = useState<string>('');
    const [tempElevenLabsKeyInput, setTempElevenLabsKeyInput] = useState('');
    const [tempWellSaidKeyInput, setTempWellSaidKeyInput] = useState('');
    const [voices, setVoices] = useState<Voice[]>([]);
    const [isFetchingVoices, setIsFetchingVoices] = useState<boolean>(false);
    const [voiceScripts, setVoiceScripts] = useState<VoiceScript[]>([]);
    const [soundEffectPrompt, setSoundEffectPrompt] = useState('');
    const [soundEffectDuration, setSoundEffectDuration] = useState<number>(4.5);
    const [soundEffectJobs, setSoundEffectJobs] = useState<SoundEffectJob[]>([]);
    const [isGeneratingSfx, setIsGeneratingSfx] = useState(false);

    // AI Script Tuner State
    const [scriptTunerInput, setScriptTunerInput] = useState<string>('');
    const [scriptTunerOutput, setScriptTunerOutput] = useState<string>('');
    const [selectedTone, setSelectedTone] = useState<string>('Documentary');
    const [customTone, setCustomTone] = useState<string>('');
    const [isTuningScript, setIsTuningScript] = useState<boolean>(false);
    const [isCopied, setIsCopied] = useState<boolean>(false);

    // Bulk Script Writer State
    const [bulkScriptMode, setBulkScriptMode] = useState<'manual' | 'auto-outline'>('manual');
    const [bulkScriptStep, setBulkScriptStep] = useState<'input' | 'review'>('input');
    const [bulkScriptTitles, setBulkScriptTitles] = useState<string>('');
    const [bulkScriptGuides, setBulkScriptGuides] = useState<string>('');
    const [bulkScriptOutline, setBulkScriptOutline] = useState<string>('');
    const [bulkScriptWordCount, setBulkScriptWordCount] = useState<number>(500);
    const [bulkScriptJobs, setBulkScriptJobs] = useState<BulkScriptJob[]>([]);
    const [bulkScriptApiProvider, setBulkScriptApiProvider] = useState<BulkScriptApiProvider>('gemini');
    const [customScriptApiKeys, setCustomScriptApiKeys] = useState<Record<string, string>>({});
    const [tempBulkScriptApiKey, setTempBulkScriptApiKey] = useState('');
    const [bulkScriptModels, setBulkScriptModels] = useState<string[]>([]);
    const [bulkScriptSelectedModel, setBulkScriptSelectedModel] = useState<string>('');
    const [isFetchingBulkScriptModels, setIsFetchingBulkScriptModels] = useState(false);
    const [bulkScriptModelError, setBulkScriptModelError] = useState<string | null>(null);

    
    const finalAudioFiles = voiceScripts.filter(s => s.audioUrl);
    const isUsingCustomImageKey = !!customImageApiKeys[apiProvider];
    const isWorkingOnBulkScripts = useMemo(() => bulkScriptJobs.some(j => j.status === 'analyzing' || j.status === 'generating'), [bulkScriptJobs]);
    const isUsingVoiceApiKey = voiceProvider === 'elevenlabs' ? !!elevenLabsApiKey : !!wellSaidApiKey;
    const hasAnyCustomKey = useMemo(() => 
        Object.values(customImageApiKeys).some(k => k) || 
        !!elevenLabsApiKey || 
        !!wellSaidApiKey ||
        Object.values(customScriptApiKeys).some(k => k),
    [customImageApiKeys, elevenLabsApiKey, wellSaidApiKey, customScriptApiKeys]);


    // --- Derived State ---
    const activeResults = useMemo(() => {
        switch (mode) {
            case 'story-weaver': return storyWeaverResults;
            case 'bulk': return bulkResults;
            case 'script-to-scene': return scriptToSceneResults;
            default: return [];
        }
    }, [mode, storyWeaverResults, bulkResults, scriptToSceneResults]);

    const setActiveResults = useMemo(() => {
        switch (mode) {
            case 'story-weaver': return setStoryWeaverResults;
            case 'bulk': return setBulkResults;
            case 'script-to-scene': return setScriptToSceneResults;
            default: return () => {}; // No-op for other modes
        }
    }, [mode]);

    const isDoneGenerating = useMemo(() => {
        const currentIsGenerating = isGenerating || isWeaving || isGeneratingScenes;
        return !currentIsGenerating && activeResults.length > 0 && activeResults.every(r => r.imageUrl || r.error);
    }, [isGenerating, isWeaving, isGeneratingScenes, activeResults]);


    // Credit system initialization
    useEffect(() => {
        const creditDataString = localStorage.getItem('dailyCreditData');
        const now = new Date();
        let resetTime = null;

        if (creditDataString) {
            try {
                const creditData = JSON.parse(creditDataString);
                const lastReset = new Date(creditData.resetTimestamp);
                
                if (now > lastReset) {
                    // Time to reset
                    setCredits(DAILY_CREDITS);
                    resetTime = new Date(now.getTime() + 24 * 60 * 60 * 1000);
                    localStorage.setItem('dailyCreditData', JSON.stringify({ count: DAILY_CREDITS, resetTimestamp: resetTime.toISOString() }));
                } else {
                    // Still within the 24-hour window
                    setCredits(creditData.count);
                    resetTime = lastReset;
                }
            } catch (e) {
                 // Corrupted data, reset
                resetTime = new Date(now.getTime() + 24 * 60 * 60 * 1000);
                localStorage.setItem('dailyCreditData', JSON.stringify({ count: DAILY_CREDITS, resetTimestamp: resetTime.toISOString() }));
                setCredits(DAILY_CREDITS);
            }
        } else {
            // First time user, initialize
            resetTime = new Date(now.getTime() + 24 * 60 * 60 * 1000);
            localStorage.setItem('dailyCreditData', JSON.stringify({ count: DAILY_CREDITS, resetTimestamp: resetTime.toISOString() }));
            setCredits(DAILY_CREDITS);
        }
        setCreditResetTime(resetTime);
    }, []);

    useEffect(() => {
        const savedTheme = localStorage.getItem('theme') as Theme | null;
        const prefersDark = window.matchMedia('(prefers-color-scheme: dark)').matches;
        const initialTheme = savedTheme || (prefersDark ? 'dark' : 'light');
        setTheme(initialTheme);
        
        // Load all keys from localStorage
        const savedImageKeys = localStorage.getItem('customImageApiKeys');
        if (savedImageKeys) {
            try {
                setCustomImageApiKeys(JSON.parse(savedImageKeys));
            } catch (e) { console.error("Could not parse image API keys from localStorage");}
        }

        const savedElevenLabsKey = localStorage.getItem('elevenLabsApiKey');
        if (savedElevenLabsKey) {
            setElevenLabsApiKey(savedElevenLabsKey);
            setTempElevenLabsKeyInput(savedElevenLabsKey);
        }

        const savedWellSaidKey = localStorage.getItem('wellSaidApiKey');
        if (savedWellSaidKey) {
            setWellSaidApiKey(savedWellSaidKey);
            setTempWellSaidKeyInput(savedWellSaidKey);
        }
        
        const savedScriptProvider = localStorage.getItem('bulkScriptApiProvider') as BulkScriptApiProvider;
        if(savedScriptProvider && BULK_SCRIPT_PROVIDERS.some(p => p.id === savedScriptProvider)) {
            setBulkScriptApiProvider(savedScriptProvider);
        }
        
        const savedScriptKeys: Record<string, string> = {};
        ['openai', 'groq', 'deepseek'].forEach(provider => {
            const key = localStorage.getItem(`bulk_script_api_key_${provider}`);
            if(key) savedScriptKeys[provider] = key;
        });
        setCustomScriptApiKeys(savedScriptKeys);


    }, []);

    // Effect to fetch styles when provider or keys change
    useEffect(() => {
        const fetchStyles = async () => {
            const service = imageServices[apiProvider];
            const apiKey = customImageApiKeys[apiProvider];
            
            // For providers that require a key to fetch styles
            const keyIsRequired = ['leonardo', 'ideogram', 'runway'].includes(apiProvider);

            if (keyIsRequired && !apiKey) {
                setImageStyles([]); // No key, no styles
                setImageStyle('None');
                return;
            }

            try {
                const styles = await service.getStyles(apiKey);
                setImageStyles(styles);
                // Reset to 'None' if current style isn't in new list, otherwise keep it
                if (!styles.includes(imageStyle)) {
                    setImageStyle(styles.length > 0 ? styles[0] : 'None');
                }
            } catch (error) {
                console.error(`Failed to fetch styles for ${apiProvider}:`, error);
                setError(`Could not load styles for ${apiProvider}.`);
                setImageStyles(DEFAULT_IMAGE_STYLES); // Fallback to default
                setImageStyle('Cinematic');
            }
        };

        fetchStyles();
    }, [apiProvider, customImageApiKeys]);


    // Effect to fetch voices when provider or key changes
    useEffect(() => {
        const fetchProviderVoices = async () => {
            setIsFetchingVoices(true);
            setError(null);
            setVoices([]);

            try {
                let fetchedVoices: Voice[] = [];
                if (voiceProvider === 'elevenlabs' && elevenLabsApiKey) {
                    fetchedVoices = await elevenLabsService.fetchVoices(elevenLabsApiKey);
                } else if (voiceProvider === 'wellsaid' && wellSaidApiKey) {
                    fetchedVoices = await wellSaidLabsService.fetchVoices(wellSaidApiKey);
                }
                setVoices(fetchedVoices);

                // Set default voice for new scripts if one isn't set, or if provider changed
                setVoiceScripts(prev => prev.map(s => {
                    const voiceIsFromCurrentProvider = fetchedVoices.some(v => v.id === s.voiceId);
                    return (s.voiceId && voiceIsFromCurrentProvider) ? s : {...s, voiceId: fetchedVoices[0]?.id || '', provider: voiceProvider };
                }));
                
            } catch (err) {
                const providerName = voiceProvider === 'elevenlabs' ? 'ElevenLabs' : 'WellSaid Labs';
                setError(`Could not fetch ${providerName} voices: ${err.message}. Please check your key.`);
                setVoices([]);
            } finally {
                setIsFetchingVoices(false);
            }
        };
        
        fetchProviderVoices();

    }, [voiceProvider, elevenLabsApiKey, wellSaidApiKey]);

    // Cleanup audio preview on unmount
    useEffect(() => {
        return () => {
            if (audioPreview) {
                audioPreview.pause();
                URL.revokeObjectURL(audioPreview.src);
            }
        };
    }, [audioPreview]);

    // Initialize with one script card for voice generator
    useEffect(() => {
        if (mode === 'voice-generator' && voiceScripts.length === 0) {
            handleAddVoiceScript();
        }
    }, [mode]);

    // Reset Bulk Script Writer when mode changes
    useEffect(() => {
        if (mode !== 'bulk-script-writer') {
            setBulkScriptJobs([]);
            setBulkScriptStep('input');
        }
    }, [mode]);

    const handleFetchAndSetModels = useCallback(async (provider: BulkScriptApiProvider, key: string) => {
        if (provider === 'gemini' || !key) {
            setBulkScriptModels([]);
            setBulkScriptSelectedModel('');
            return;
        }
    
        setIsFetchingBulkScriptModels(true);
        setBulkScriptModelError(null);
        setBulkScriptModels([]); // Clear old models
    
        try {
            const models = await multiProviderService.fetchModelsForProvider(provider, key);
            setBulkScriptModels(models);
    
            const savedModel = localStorage.getItem(`bulk_script_model_${provider}`);
            // Try to find a good default model, falling back to the first in the list
            const recommendedModel = models.find(m => m.includes('llama3-8b')) || models.find(m => m.includes('mixtral')) || models[0];
            
            if (savedModel && models.includes(savedModel)) {
                setBulkScriptSelectedModel(savedModel);
            } else if (recommendedModel) {
                setBulkScriptSelectedModel(recommendedModel);
                localStorage.setItem(`bulk_script_model_${provider}`, recommendedModel);
            } else {
                setBulkScriptSelectedModel('');
            }
    
        } catch (err) {
            const message = err instanceof Error ? err.message : "Failed to fetch models. Check the API key and provider.";
            setBulkScriptModelError(message);
            setBulkScriptModels([]);
            setBulkScriptSelectedModel('');
        } finally {
            setIsFetchingBulkScriptModels(false);
        }
    }, []);

    // Effect to handle provider changes for bulk script writer
    useEffect(() => {
        const key = customScriptApiKeys[bulkScriptApiProvider as keyof typeof customScriptApiKeys] || '';
        setTempBulkScriptApiKey(key);

        const savedModel = localStorage.getItem(`bulk_script_model_${bulkScriptApiProvider}`);
        setBulkScriptSelectedModel(savedModel || '');
    
        if (key && bulkScriptApiProvider !== 'gemini') {
            handleFetchAndSetModels(bulkScriptApiProvider, key);
        } else {
            setBulkScriptModels([]);
            setBulkScriptModelError(null);
        }
    }, [bulkScriptApiProvider, customScriptApiKeys, handleFetchAndSetModels]);
    
    // Effect to sync temp image API key when provider changes
    useEffect(() => {
        const key = customImageApiKeys[apiProvider as keyof typeof customImageApiKeys] || '';
        setTempImageApiKey(key);
    }, [apiProvider, customImageApiKeys]);

    const spendCredits = useCallback((amount: number): boolean => {
        // Credits only apply to default Gemini provider without a custom key
        if (apiProvider !== 'gemini' || !!customImageApiKeys.gemini) return true;
    
        if (credits < amount) {
            setError(`Not enough credits. Required: ${amount}, Available: ${credits}. Credits reset daily.`);
            return false;
        }
        
        const newCreditCount = credits - amount;
        setCredits(newCreditCount);
        
        if (creditResetTime) {
            const creditData = {
                count: newCreditCount,
                resetTimestamp: creditResetTime.toISOString()
            };
            localStorage.setItem('dailyCreditData', JSON.stringify(creditData));
        }
        return true;
    }, [credits, creditResetTime, apiProvider, customImageApiKeys]);


    const handleSaveSingleKey = (section: keyof AllApiKeys, provider: string, key: string) => {
        const trimmedKey = key.trim();
    
        switch (section) {
            case 'image': {
                const newKeys = { ...customImageApiKeys, [provider]: trimmedKey };
                setCustomImageApiKeys(newKeys);
                localStorage.setItem('customImageApiKeys', JSON.stringify(newKeys));
                break;
            }
            case 'voice': {
                if (provider === 'elevenlabs') {
                    setElevenLabsApiKey(trimmedKey);
                    setTempElevenLabsKeyInput(trimmedKey);
                    localStorage.setItem('elevenLabsApiKey', trimmedKey);
                } else if (provider === 'wellsaid') {
                    setWellSaidApiKey(trimmedKey);
                    setTempWellSaidKeyInput(trimmedKey);
                    localStorage.setItem('wellSaidApiKey', trimmedKey);
                }
                break;
            }
            case 'script': {
                const newKeys = { ...customScriptApiKeys, [provider]: trimmedKey };
                setCustomScriptApiKeys(newKeys);
                if (trimmedKey) {
                    localStorage.setItem(`bulk_script_api_key_${provider}`, trimmedKey);
                } else {
                    localStorage.removeItem(`bulk_script_api_key_${provider}`);
                }
                
                // Also sync the temp key if it's the current provider
                if (provider === bulkScriptApiProvider) {
                    setTempBulkScriptApiKey(trimmedKey);
                }
                break;
            }
        }
    };

    const handleSaveImageApiKey = () => {
        handleSaveSingleKey('image', apiProvider, tempImageApiKey);
        alert('API Key saved!');
    };

    const handleSaveAllApiKeys = (keys: AllApiKeys) => {
        // Image Keys
        setCustomImageApiKeys(keys.image);
        localStorage.setItem('customImageApiKeys', JSON.stringify(keys.image));

        // Voice Keys
        const elevenKey = keys.voice.elevenlabs || '';
        setElevenLabsApiKey(elevenKey);
        localStorage.setItem('elevenLabsApiKey', elevenKey);
        setTempElevenLabsKeyInput(elevenKey);

        const wellSaidKey = keys.voice.wellsaid || '';
        setWellSaidApiKey(wellSaidKey);
        localStorage.setItem('wellSaidApiKey', wellSaidKey);
        setTempWellSaidKeyInput(wellSaidKey);
        
        // Script Keys
        const scriptKeys = keys.script || {};
        setCustomScriptApiKeys(scriptKeys);
        ['openai', 'groq', 'deepseek'].forEach(provider => {
            const key = scriptKeys[provider as keyof typeof scriptKeys] || '';
            if (key) {
                localStorage.setItem(`bulk_script_api_key_${provider}`, key);
            } else {
                localStorage.removeItem(`bulk_script_api_key_${provider}`);
            }
        });

        // Also sync the temp key for the currently selected bulk script provider
        const currentScriptKey = scriptKeys[bulkScriptApiProvider as keyof typeof scriptKeys];
        if (currentScriptKey !== undefined) {
             setTempBulkScriptApiKey(currentScriptKey);
        }
        
        setError(null);
    };

    const handleClearAllApiKeys = () => {
        // Clear Image
        setCustomImageApiKeys({});
        localStorage.removeItem('customImageApiKeys');

        // Clear Voice
        setElevenLabsApiKey('');
        localStorage.removeItem('elevenLabsApiKey');
        setTempElevenLabsKeyInput('');
        setWellSaidApiKey('');
        localStorage.removeItem('wellSaidApiKey');
        setTempWellSaidKeyInput('');
        setVoices([]);
        
        // Clear Scripts
        setCustomScriptApiKeys({});
        setTempBulkScriptApiKey('');
        ['openai', 'groq', 'deepseek'].forEach(provider => {
            localStorage.removeItem(`bulk_script_api_key_${provider}`);
        });

        setError(null);
    };

    useEffect(() => {
        if (theme === 'dark') {
            document.documentElement.classList.add('dark');
        } else {
            document.documentElement.classList.remove('dark');
        }
        localStorage.setItem('theme', theme);
    }, [theme]);
    
    const handleStartCreating = () => {
        setHasStarted(true);
        setMode('story-weaver');
        setIsPanelOpen(true);
    };

    const handleNavClick = (newMode: Mode) => {
        setHasStarted(true);
        if (mode === newMode && isPanelOpen) {
            setIsPanelOpen(false);
        } else {
            setMode(newMode);
            setBulkScriptStep('input'); // Always reset step when switching modes
            setIsPanelOpen(true);
        }
    };

    const setVideoConfig = useCallback((updater: (prevState: VideoConfig) => VideoConfig) => {
        setVideoConfigState(currentState => {
            const newPresent = updater(currentState.present);
            if (JSON.stringify(newPresent) === JSON.stringify(currentState.present)) {
                return currentState; // No change
            }
            const newPast = [...currentState.past, currentState.present];
            // Limit history size
            if (newPast.length > 30) {
                newPast.shift();
            }
            return {
                past: newPast,
                present: newPresent,
                future: [] // Clear future on new action
            };
        });
    }, []);

    const undoVideoConfig = useCallback(() => {
        setVideoConfigState(currentState => {
            const { past, present, future } = currentState;
            if (past.length === 0) return currentState;

            const previous = past[past.length - 1];
            const newPast = past.slice(0, past.length - 1);
            
            return {
                past: newPast,
                present: previous,
                future: [present, ...future]
            };
        });
    }, []);

    const redoVideoConfig = useCallback(() => {
        setVideoConfigState(currentState => {
            const { past, present, future } = currentState;
            if (future.length === 0) return currentState;

            const next = future[0];
            const newFuture = future.slice(1);
            
            return {
                past: [...past, present],
                present: next,
                future: newFuture
            };
        });
    }, []);

    const handlePreviewVoice = async (voiceId: string, provider: VoiceProvider) => {
        if (audioPreview) {
            audioPreview.pause();
        }

        if (previewStatus.id === voiceId && previewStatus.state !== 'loading') {
            setPreviewStatus({ id: null, state: 'idle' });
            setAudioPreview(null);
            return;
        }
        
        const apiKey = provider === 'elevenlabs' ? elevenLabsApiKey : wellSaidApiKey;
        if (!apiKey) {
            const providerName = provider === 'elevenlabs' ? 'ElevenLabs' : 'WellSaid Labs';
            setError(`Please set your ${providerName} API key to preview voices.`);
            return;
        }

        setPreviewStatus({ id: voiceId, state: 'loading' });
        setError(null);

        try {
            const sampleText = "You can use this voice to create audio for your content.";
            let audioBlob: Blob;

            if (provider === 'elevenlabs') {
                audioBlob = await elevenLabsService.generateAudio(apiKey, voiceId, sampleText, 50);
            } else { // wellsaid
                audioBlob = await wellSaidLabsService.generateAudio(apiKey, voiceId, sampleText);
            }

            const audioUrl = URL.createObjectURL(audioBlob);
            const newAudio = new Audio(audioUrl);
            setAudioPreview(newAudio);
            
            newAudio.play();
            
            newAudio.onplaying = () => {
                 setPreviewStatus({ id: voiceId, state: 'playing' });
            };
            newAudio.onended = () => {
                setPreviewStatus({ id: null, state: 'idle' });
                setAudioPreview(null);
                URL.revokeObjectURL(audioUrl);
            };
            newAudio.onerror = () => {
                setError("Failed to play audio preview.");
                setPreviewStatus({ id: null, state: 'idle' });
                setAudioPreview(null);
                URL.revokeObjectURL(audioUrl);
            };

        } catch (e) {
            const errorMessage = e instanceof Error ? e.message : 'An unknown error occurred';
            setError(`Preview failed: ${errorMessage}`);
            setPreviewStatus({ id: null, state: 'idle' });
        }
    };

    const blobToBase64 = (blob: Blob): Promise<string> => new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onloadend = () => resolve((reader.result as string).split(',')[1]);
        reader.onerror = reject;
        reader.readAsDataURL(blob);
    });

    const handleAudioFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (file) {
            setAudioBlob(file);
        }
        if (e.target) e.target.value = '';
    };

    const handleTranscribe = async () => {
        if (!audioBlob) return;
        setIsTranscribing(true);
        setError(null);
        setWeavingStep('Transcribing...');

        try {
            const base64Audio = await blobToBase64(audioBlob);
            const script = await geminiService.transcribeAudio(base64Audio, audioBlob.type);
            setScriptInput(script);
            setStoryWeaverMode('script'); // Switch to script mode for editing
        } catch(e) {
            setError(e instanceof Error ? e.message : 'Transcription failed.');
        } finally {
            setIsTranscribing(false);
            setAudioBlob(null);
            setWeavingStep('');
        }
    };

    const handleFileChange = async (e: React.ChangeEvent<HTMLInputElement>, setter: (value: string) => void, id?: string) => {
        const file = e.target.files?.[0];
        if (!file) return;

        try {
            let text;
            if (file.name.endsWith('.docx')) {
                const arrayBuffer = await file.arrayBuffer();
                const result = await mammoth.extractRawText({ arrayBuffer });
                text = result.value;
            } else {
                text = await file.text();
            }
            if (id) {
                // This is for voice script
                handleUpdateVoiceScript(id, { text });
            } else {
                setter(text);
            }
        } catch (error) {
            console.error("Error reading file:", error);
            setError("Could not read file. Please ensure it's a valid .txt or .docx file.");
        } finally {
           if (e.target) e.target.value = '';
        }
    };

    const handleStyleReferenceChange = (e: React.ChangeEvent<HTMLInputElement>) => {
        const file = e.target.files?.[0];
        if (!file) return;
    
        const reader = new FileReader();
        reader.onloadend = () => {
            const url = reader.result as string;
            const base64 = url.split(',')[1];
            setStyleReferenceImage({ url, base64 });
        };
        reader.readAsDataURL(file);
    
        if (e.target) e.target.value = ''; // Reset file input
    };
    
    const clearStyleReference = () => {
        setStyleReferenceImage(null);
        if(styleReferenceFileInputRef.current) {
            styleReferenceFileInputRef.current.value = '';
        }
    };

    const processAndGenerate = async (promptsList: string[]) => {
        if (promptsList.length === 0) return;
        setHasStarted(true);
        setError(null);
        setIsGenerating(true);
        setVideoUrl(null);
        setProgress(0);
        
        const initialJobs: ImageJob[] = promptsList.map((prompt, index) => ({
            id: `${Date.now()}-${index}`,
            prompt,
        }));
        setActiveResults(initialJobs);

        const generateFn = imageServices[apiProvider].generateImages;
        const apiKey = customImageApiKeys[apiProvider] || null;
        
        await generateFn({
            prompts: promptsList,
            ratio: ratio,
            steps: steps,
            apiKey: apiKey,
            onProgress: (index, result) => {
                setActiveResults(prevResults => {
                    const newResults = [...prevResults];
                    if (newResults[index]) {
                       newResults[index] = { ...newResults[index], imageUrl: result.url, error: result.error };
                    }
                    return newResults;
                });
                setProgress(p => p + 1);
            }
        });

        setIsGenerating(false);
    };

    const handleGenerateClick = async () => {
        const promptsList = promptsInput.split('\n').map(p => p.trim()).filter(p => p);
        if (promptsList.length === 0) return;

        if (!spendCredits(promptsList.length)) return;

        let finalPrompts = imageStyle !== 'None'
            ? promptsList.map(p => `${p}, in a ${imageStyle.toLowerCase()} style`)
            : promptsList;

        try {
            if (apiProvider === 'gemini' && styleReferenceImage) {
                setIsDescribingStyle(true);
                setError(null);
                const styleDescription = await geminiService.describeImageStyle(styleReferenceImage.base64);
                finalPrompts = finalPrompts.map(p => `${styleDescription}, ${p}`);
                setIsDescribingStyle(false);
            }
            await processAndGenerate(finalPrompts);
        } catch (e) {
            setError(e instanceof Error ? `Failed to process request: ${e.message}` : 'An unknown error occurred.');
            setIsGenerating(false);
            setIsDescribingStyle(false);
        }
    };
    
    const handleAnalyzeScript = async () => {
        if (!scriptInput.trim() || isAnalyzingScript) return;
        setIsAnalyzingScript(true);
        setError(null);
        setExtractedCharacters([]);
        setMainCharacter(null);
        setCharacterDetails({});
        try {
            const { characters, mainCharacter } = await geminiService.extractCharactersFromScript(scriptInput);
            setExtractedCharacters(characters);
            setMainCharacter(mainCharacter);
            const initialDetails: Record<string, CharacterDetails> = {};
            characters.forEach(char => {
                initialDetails[char] = { description: '', refImage: null };
            });
            setCharacterDetails(initialDetails);
        } catch (e) {
            setError(e instanceof Error ? e.message : 'Failed to analyze script for characters.');
        } finally {
            setIsAnalyzingScript(false);
        }
    };

    const handleCharacterRefImageUpload = async (e: React.ChangeEvent<HTMLInputElement>, charName: string) => {
        const file = e.target.files?.[0];
        if (!file) return;

        setIsDescribingImage(charName);
        setError(null);
        
        const reader = new FileReader();
        reader.onloadend = async () => {
            try {
                const url = reader.result as string;
                const base64 = url.split(',')[1];
                const newRefImage = { url, base64 };
                
                // Set image first for immediate feedback
                setCharacterDetails(prev => ({
                    ...prev,
                    [charName]: { ...(prev[charName] || {description: ''}), refImage: newRefImage }
                }));

                const description = await geminiService.describeImage(base64);
                setCharacterDetails(prev => ({
                    ...prev,
                    [charName]: { ...(prev[charName] || {refImage: null}), description }
                }));
            } catch (e) {
                setError(`Failed to get description for ${charName}: ${e instanceof Error ? e.message : 'Unknown error'}`);
                // Revert image if description fails
                 setCharacterDetails(prev => ({
                    ...prev,
                    [charName]: { ...(prev[charName] || {description: ''}), refImage: null }
                }));
            } finally {
                setIsDescribingImage(null);
            }
        };
        reader.readAsDataURL(file);

        if (e.target) e.target.value = '';
    };

    const handleScriptToSceneGenerate = async () => {
        if (!scriptInput.trim()) return;

        setIsGeneratingScenes(true);
        setError(null);

        try {
            const finalCharacterDetails: Record<string, { description: string }> = {};
            Object.entries(characterDetails).forEach(([name, details]) => {
                if (details.description.trim()) {
                    finalCharacterDetails[name] = { description: details.description };
                }
            });

            const scenePrompts = await geminiService.generateScenesFromScript(scriptInput, numScenes, finalCharacterDetails);

            if (!spendCredits(scenePrompts.length)) {
                setIsGeneratingScenes(false);
                return;
            }

            let finalPrompts = imageStyle !== 'None'
                ? scenePrompts.map(p => `${p}, in a ${imageStyle.toLowerCase()} style`)
                : scenePrompts;

            if (apiProvider === 'gemini' && styleReferenceImage) {
                setIsDescribingStyle(true);
                const styleDescription = await geminiService.describeImageStyle(styleReferenceImage.base64);
                finalPrompts = finalPrompts.map(p => `${styleDescription}, ${p}`);
                setIsDescribingStyle(false);
            }

            await processAndGenerate(finalPrompts);

        } catch (e) {
            setError(e instanceof Error ? e.message : 'Failed to generate scenes.');
        } finally {
            setIsGeneratingScenes(false);
            setIsDescribingStyle(false);
        }
    };

    const handleWeaveStory = async () => {
        if (isWeaving) return;

        setIsWeaving(true);
        setError(null);
        setStoryWeaverResults([]);
        setScriptInput(storyWeaverMode === 'idea' ? '' : scriptInput);
        setVideoUrl(null);

        try {
            let scriptToProcess: string;
            let scenesToGenerate: number = numScenes;

            if (storyWeaverMode === 'idea') {
                if (!storyIdea.trim()) throw new Error("Please provide a story idea.");
                setWeavingStep('script');
                scriptToProcess = await geminiService.generateScriptFromIdea(storyIdea, videoDuration);
                scenesToGenerate = Math.max(1, Math.round(videoDuration * 5)); // Approx 5 scenes per minute
                setNumScenes(scenesToGenerate); // Update state for progress display
                setScriptInput(scriptToProcess); // Show the generated script in the UI
            } else { // 'script' mode
                if (!scriptInput.trim()) throw new Error("Please provide a script.");
                scriptToProcess = scriptInput;
            }

            setWeavingStep('scenes');
            const characterDetailsForWeaving: Record<string, { description: string }> = {};
            if (activeCharacter?.prompt) {
                characterDetailsForWeaving['MAIN_CHARACTER'] = { description: activeCharacter.prompt };
            }

            const scenePrompts = await geminiService.generateScenesFromScript(scriptToProcess, scenesToGenerate, characterDetailsForWeaving);

            if (!spendCredits(scenePrompts.length)) {
                throw new Error(`Not enough credits. Required: ${scenePrompts.length}, Available: ${credits}.`);
            }

            setWeavingStep('images');
            let finalPrompts = scenePrompts.map(p => `${p}, in a ${imageStyle.toLowerCase()} style`);

            if (apiProvider === 'gemini' && styleReferenceImage) {
                setIsDescribingStyle(true);
                const styleDescription = await geminiService.describeImageStyle(styleReferenceImage.base64);
                finalPrompts = finalPrompts.map(p => `${styleDescription}, ${p}`);
                setIsDescribingStyle(false);
            }

            await processAndGenerate(finalPrompts);

        } catch (e) {
            setError(e instanceof Error ? e.message : 'Storyboard process failed.');
        } finally {
            setIsWeaving(false);
            setWeavingStep('');
            setIsDescribingStyle(false);
        }
    };

    const generateCharacterSheet = async () => {
        if (!characterDescription.trim()) return;
        
        const prompts = [
            `Full-body character sheet for a character described as: ${characterDescription}. Multiple poses and expressions, neutral background, consistent design, front view`,
            `Full-body character sheet for a character described as: ${characterDescription}. Multiple poses and expressions, neutral background, consistent design, side view`,
            `Full-body character sheet for a character described as: ${characterDescription}. Multiple poses and expressions, neutral background, consistent design, back view`,
            `Full-body character sheet for a character described as: ${characterDescription}. Multiple poses and expressions, neutral background, consistent design, dynamic action pose`,
        ];

        if (!spendCredits(prompts.length)) return;

        setIsGeneratingCharacter(true);
        setError(null);
        setCharacterSheet([]);
        
        const initialJobs: ImageJob[] = prompts.map((prompt, index) => ({
            id: `char-${Date.now()}-${index}`,
            prompt,
        }));
        setCharacterSheet(initialJobs);

        await geminiService.generateImages({
            prompts,
            ratio: '1:1',
            apiKey: customImageApiKeys.gemini || null,
            onProgress: (index, result) => {
                setCharacterSheet(prev => {
                    const newSheet = [...prev];
                    if (newSheet[index]) {
                       newSheet[index] = { ...newSheet[index], imageUrl: result.url, error: result.error };
                    }
                    return newSheet;
                });
            }
        });
        
        setIsGeneratingCharacter(false);
    };
    
    const handleRegenerateOne = async (jobId: string) => {
        const job = activeResults.find(r => r.id === jobId);
        if (!job) return;

        if (!spendCredits(1)) return;

        setActiveResults(prev => prev.map(r => r.id === jobId ? { ...r, imageUrl: undefined, error: undefined } : r));

        const generateFn = imageServices[apiProvider].generateImages;
        const apiKey = customImageApiKeys[apiProvider] || null;
        
        await generateFn({
            prompts: [job.prompt],
            ratio: ratio,
            steps: steps,
            apiKey: apiKey,
            onProgress: (index, result) => { // index will be 0
                setActiveResults(prev => prev.map(r => r.id === jobId ? { ...r, imageUrl: result.url, error: result.error } : r));
            }
        });
    };
    
    const handleSaveEdit = (jobId: string, newPrompt: string, newImageUrl: string) => {
        setActiveResults(prev => prev.map(r => r.id === jobId ? { ...r, prompt: newPrompt, imageUrl: newImageUrl } : r));
        setEditingJob(null);
    };

    const handleDownloadOne = (job: ImageJob) => {
        if (!job.imageUrl) return;
    
        const link = document.createElement("a");
        link.href = job.imageUrl;
        // Create a user-friendly filename from the prompt
        const sanitizedPrompt = job.prompt.substring(0, 40).replace(/[^a-zA-Z0-9]/g, '_').replace(/_{2,}/g, '_').toLowerCase();
        const fileName = `${sanitizedPrompt}_${job.id.slice(-6)}.png`;
        link.download = fileName;
        
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    };

    const downloadAll = () => {
        const zip = new JSZip();
        const imageFolder = zip.folder("images");
        const successfulResults = activeResults.filter(r => r.imageUrl && !r.error);

        if (!imageFolder || successfulResults.length === 0) return;

        let promptText = "";
        const promises = successfulResults.map((result, i) => {
            promptText += `Image ${i + 1}:\n${result.prompt}\n\n`;
            const base64Data = result.imageUrl!.split(',')[1];
            return imageFolder.file(`image_${i + 1}.png`, base64Data, { base64: true });
        });
        
        zip.file("prompts.txt", promptText);

        Promise.all(promises).then(() => {
            zip.generateAsync({ type: "blob" }).then(content => {
                const link = document.createElement("a");
                link.href = URL.createObjectURL(content);
                link.download = `ai_images_${Date.now()}.zip`;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
            });
        });
    };

    const clearAll = () => {
        setBulkResults([]);
        setScriptToSceneResults([]);
        setStoryWeaverResults([]);
        setPromptsInput('');
        setScriptInput('');
        setError(null);
        setProgress(0);
        setVideoUrl(null);
        setHasStarted(false);
        setIsPanelOpen(false);
        setExtractedCharacters([]);
        setMainCharacter(null);
        setCharacterDetails({});
        setBulkScriptJobs([]);
        setStyleReferenceImage(null);
    };

    const wrapText = (context: CanvasRenderingContext2D, text: string, maxWidth: number, wordsPerLine: number) => {
        const words = text.split(' ');
        let line = '';
        const lines = [];
    
        for(let n = 0; n < words.length; n++) {
            const testLine = line + words[n] + ' ';
            const metrics = context.measureText(testLine);
            const testWidth = metrics.width;
            const currentWordsInLine = line.trim().split(' ').filter(w => w !== '').length;

            if ((testWidth > maxWidth || currentWordsInLine >= wordsPerLine) && n > 0) {
                lines.push(line.trim());
                line = words[n] + ' ';
            } else {
                line = testLine;
            }
        }
        lines.push(line.trim());
        return lines.filter(l => l);
    };


    const handleGenerateVideo = async () => {
        setIsGeneratingVideo(true);
        setError(null);
        setVideoUrl(null);

        const validScenes = activeResults.filter(r => r.imageUrl);
        if (validScenes.length === 0) {
            setError("No valid images to generate video from.");
            setIsGeneratingVideo(false);
            return;
        }

        const [w, h] = ratio.split(':').map(Number);
        const canvasWidth = 1280;
        const canvasHeight = (canvasWidth * h) / w;

        const canvas = document.createElement('canvas');
        canvas.width = canvasWidth;
        canvas.height = canvasHeight;
        const ctx = canvas.getContext('2d');
        if (!ctx) {
            setError("Could not create canvas context.");
            setIsGeneratingVideo(false);
            return;
        }

        const hexToRgb = (hex: string) => {
            const result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
            return result ? { r: parseInt(result[1], 16), g: parseInt(result[2], 16), b: parseInt(result[3], 16) } : null;
        };

        const audioContext = new AudioContext();
        const audioDestination = audioContext.createMediaStreamDestination();

        // Step 1: Prepare all scene data (image, audio, duration, words) upfront
        const sceneDataPromises = validScenes.map(async (scene) => {
            const imagePromise = new Promise<HTMLImageElement>(resolve => {
                const img = new Image();
                img.crossOrigin = 'anonymous';
                img.src = scene.imageUrl!;
                img.onload = () => resolve(img);
            });

            const words = scene.prompt.trim().split(' ').filter(w => w.length > 0);
            let audioBuffer: AudioBuffer | null = null;
            let audioDuration = 5000; // Default 5s in ms

            const voice = voices.find(v => v.id === videoConfig.voice);
            if (voice && words.length > 0) {
                const apiKey = voice.provider === 'elevenlabs' ? elevenLabsApiKey : wellSaidApiKey;
                try {
                    let audioBlob: Blob;
                    if(voice.provider === 'elevenlabs' && elevenLabsApiKey) {
                        audioBlob = await elevenLabsService.generateAudio(elevenLabsApiKey, videoConfig.voice, scene.prompt, 50);
                    } else if (voice.provider === 'wellsaid' && wellSaidApiKey) {
                        audioBlob = await wellSaidLabsService.generateAudio(wellSaidApiKey, videoConfig.voice, scene.prompt);
                    } else {
                        audioBlob = new Blob();
                    }
                    
                    if (audioBlob.size > 0) {
                        const arrayBuffer = await audioBlob.arrayBuffer();
                        if (audioContext.state === 'suspended') await audioContext.resume();
                        const decodedBuffer = await audioContext.decodeAudioData(arrayBuffer);
                        audioBuffer = decodedBuffer;
                        audioDuration = decodedBuffer.duration * 1000;
                    }
                } catch (e) {
                    console.error(`Failed to process audio for prompt: "${scene.prompt}"`, e);
                }
            }

            const image = await imagePromise;
            return { image, audioBuffer, duration: audioDuration, prompt: scene.prompt, words };
        });

        const allSceneData = await Promise.all(sceneDataPromises);
        
        // Step 2: Create a master timeline with pre-calculated start/end times
        let cumulativeTime = 0;
        const scenesWithTimings = allSceneData.map((scene, index) => {
            const startTime = cumulativeTime;
            cumulativeTime += scene.duration;
            return { ...scene, index, startTime, endTime: cumulativeTime };
        });
        const totalVideoDuration = cumulativeTime;

        const stream = canvas.captureStream(30);
        const finalStream = new MediaStream([...stream.getTracks(), ...audioDestination.stream.getAudioTracks()]);
        const recorder = new MediaRecorder(finalStream, { mimeType: 'video/webm' });
        const chunks: Blob[] = [];
        recorder.ondataavailable = e => chunks.push(e.data);
        recorder.onstop = () => {
            try {
                const blob = new Blob(chunks, { type: 'video/webm' });
                setVideoUrl(URL.createObjectURL(blob));
                audioContext.close();
            } catch(e) {
                setError("Failed to create video file.");
                console.error(e)
            } finally {
                setIsGeneratingVideo(false);
            }
        };
        recorder.start();

        let videoStartTime = -1;
        const playedAudio = new Set<number>();

        const draw = (time: number) => {
            if (recorder.state !== 'recording') return;
            
            if (videoStartTime < 0) {
                videoStartTime = time;
            }
            const totalElapsed = time - videoStartTime;

            if (totalElapsed >= totalVideoDuration) {
                if (recorder.state === "recording") recorder.stop();
                return;
            }

            // Step 3: Find the current scene based on the master timeline
            let currentSceneData = scenesWithTimings[scenesWithTimings.length - 1]; // Default to last scene
            for(const scene of scenesWithTimings) {
                if (totalElapsed < scene.endTime) {
                    currentSceneData = scene;
                    break;
                }
            }

            const sceneElapsed = totalElapsed - currentSceneData.startTime;
            
            // Step 4: Play audio for the current scene if it hasn't been played
            if (!playedAudio.has(currentSceneData.index) && currentSceneData.audioBuffer) {
                const source = audioContext.createBufferSource();
                source.buffer = currentSceneData.audioBuffer;
                source.connect(audioDestination);
                source.start();
                playedAudio.add(currentSceneData.index);
            }
            
            ctx.clearRect(0, 0, canvasWidth, canvasHeight);
            
            // --- Draw Scene Image with Transitions ---
            const transitionDuration = 500;
            const transitionStart = currentSceneData.duration - transitionDuration;
            const nextScene = scenesWithTimings[currentSceneData.index + 1];

            if (nextScene && sceneElapsed > transitionStart && videoConfig.sceneTransition !== 'none') {
                const transProgress = (sceneElapsed - transitionStart) / transitionDuration;
                if (videoConfig.sceneTransition === 'fade') {
                    ctx.globalAlpha = 1;
                    ctx.drawImage(currentSceneData.image, 0, 0, canvasWidth, canvasHeight);
                    ctx.globalAlpha = Math.max(0, Math.min(1, transProgress));
                    ctx.drawImage(nextScene.image, 0, 0, canvasWidth, canvasHeight);
                    ctx.globalAlpha = 1;
                } else if (videoConfig.sceneTransition === 'slide-left') {
                     const progress = Math.max(0, Math.min(1, transProgress));
                     ctx.drawImage(currentSceneData.image, -progress * canvasWidth, 0, canvasWidth, canvasHeight);
                     ctx.drawImage(nextScene.image, canvasWidth - progress * canvasWidth, 0, canvasWidth, canvasHeight);
                }
            } else {
                ctx.drawImage(currentSceneData.image, 0, 0, canvasWidth, canvasHeight);
            }

            // --- Draw Captions ---
            if (videoConfig.captionsEnabled && currentSceneData.words.length > 0) {
                const durationPerWord = currentSceneData.duration / currentSceneData.words.length;
                const wordsToShow = Math.floor(sceneElapsed / durationPerWord) + 1;
                const currentText = currentSceneData.words.slice(0, wordsToShow).join(' ');

                ctx.font = `bold 36px "${videoConfig.captionFont}", "Inter", sans-serif`;
                ctx.textAlign = 'center';

                if(videoConfig.shadowEnabled) {
                    ctx.shadowColor = videoConfig.captionShadowColor;
                    ctx.shadowBlur = videoConfig.captionShadowBlur;
                }
                
                if(videoConfig.strokeEnabled) {
                    ctx.strokeStyle = videoConfig.captionStrokeColor;
                    ctx.lineWidth = videoConfig.captionStrokeWidth;
                }
                
                const lines = wrapText(ctx, currentText, canvasWidth * 0.9, videoConfig.captionWordsPerLine);
                const totalTextHeight = lines.length * 48;
                
                const fadeInDuration = 500;
                const fadeOutStart = currentSceneData.duration - 500;
                let animProgress = 1;

                if (sceneElapsed < fadeInDuration) {
                    animProgress = sceneElapsed / fadeInDuration;
                } else if (sceneElapsed > fadeOutStart) {
                    animProgress = 1 - ((sceneElapsed - fadeOutStart) / 500);
                }
                animProgress = Math.max(0, Math.min(1, animProgress));

                if (videoConfig.captionAnimation === 'fade-in' || videoConfig.captionAnimation === 'slide-up' || videoConfig.captionAnimation === 'pop-in') {
                    ctx.globalAlpha = animProgress;
                }
                
                lines.forEach((line, i) => {
                    const yPos = canvasHeight - totalTextHeight + (i * 48) - 40;
                    let yOffset = 0;
                    let scale = 1;
                    
                    if(videoConfig.captionAnimation === 'slide-up') yOffset = (1 - animProgress) * 50;
                    if(videoConfig.captionAnimation === 'pop-in') scale = 0.95 + animProgress * 0.05;
                    
                    const lineX = canvasWidth / 2;
                    const lineY = yPos + yOffset;

                    ctx.save();
                    ctx.translate(lineX, lineY);
                    ctx.scale(scale, scale);
                    
                    if (videoConfig.backgroundEnabled) {
                        const metrics = ctx.measureText(line);
                        const rgb = hexToRgb(videoConfig.captionBackgroundColor);
                        ctx.fillStyle = rgb ? `rgba(${rgb.r}, ${rgb.g}, ${rgb.b}, ${videoConfig.captionBackgroundOpacity * animProgress})` : 'rgba(0,0,0,0.5)';
                        
                        ctx.fillRect(-(metrics.width / 2) - 10, -38, metrics.width + 20, 48);
                    }
                    
                    ctx.fillStyle = videoConfig.captionColor;
                    if(videoConfig.strokeEnabled) ctx.strokeText(line, 0, 0);
                    ctx.fillText(line, 0, 0);
                    ctx.restore();
                });

                ctx.globalAlpha = 1;
                ctx.shadowBlur = 0;
                ctx.shadowColor = 'transparent';
                ctx.lineWidth = 1;
            }
            
            requestAnimationFrame(draw);
        };

        requestAnimationFrame(draw);
    };

    // --- AI Voice Generator Functions ---
    const handleAddVoiceScript = () => {
        setVoiceScripts(prev => [...prev, {
            id: `vs-${Date.now()}`,
            text: '',
            voiceId: voices[0]?.id || '',
            provider: voiceProvider,
            speed: 50, // Corresponds to 100% rate
        }]);
    };

    const handleRemoveVoiceScript = (id: string) => {
        setVoiceScripts(prev => prev.filter(s => s.id !== id));
    };

    const handleUpdateVoiceScript = (id: string, updates: Partial<VoiceScript>) => {
        setVoiceScripts(prev => prev.map(s => {
            if (s.id === id) {
                const isJustAnimationUpdate = Object.keys(updates).length === 1 && 'justUpdated' in updates;
                const newScript = { ...s, ...updates };
    
                if (!isJustAnimationUpdate) {
                    newScript.error = undefined; // Clear error on any "real" update
                }
    
                // If text, voice, or speed is changed AND audio already exists, mark as modified.
                if (newScript.audioUrl && ('text' in updates || 'voiceId' in updates || 'speed' in updates)) {
                    newScript.isModified = true;
                }
                return newScript;
            }
            return s;
        }));
    };
    
    const handleGenerateSingleVoice = async (id: string, isBulk = false) => {
        const script = voiceScripts.find(s => s.id === id);
        if (!script || !script.text.trim() || !script.voiceId) {
            if (!isBulk) handleUpdateVoiceScript(id, { error: 'Script text and voice character must be set.' });
            return;
        }

        const apiKey = script.provider === 'elevenlabs' ? elevenLabsApiKey : wellSaidApiKey;
        if (!apiKey) {
            const providerName = script.provider === 'elevenlabs' ? 'ElevenLabs' : 'WellSaid Labs';
            if (!isBulk) handleUpdateVoiceScript(id, { error: `${providerName} API Key is not set.` });
            return;
        }
        
        handleUpdateVoiceScript(id, { isGenerating: true, error: undefined });

        try {
            let audioBlob: Blob;
            if (script.provider === 'elevenlabs') {
                audioBlob = await elevenLabsService.generateAudio(apiKey, script.voiceId, script.text, script.speed);
            } else { // 'wellsaid'
                audioBlob = await wellSaidLabsService.generateAudio(apiKey, script.voiceId, script.text);
            }

            if (audioBlob.size === 0) {
                throw new Error("Generation resulted in an empty audio file. The script may be too short.");
            }
            const audioUrl = URL.createObjectURL(audioBlob);
            handleUpdateVoiceScript(id, { audioUrl, isGenerating: false, isModified: false, justUpdated: true });

            // Remove the visual cue after the animation duration
            setTimeout(() => {
                handleUpdateVoiceScript(id, { justUpdated: false });
            }, 2000);

        } catch (e) {
            const error = e instanceof Error ? e.message : 'An unknown error occurred.';
            handleUpdateVoiceScript(id, { error, isGenerating: false });
        }
    };

    const handleGenerateAllVoices = async () => {
        const validScripts = voiceScripts.filter(s => s.text.trim() && s.voiceId);
        if (validScripts.length === 0 || !isUsingVoiceApiKey) return;
        
        for (const script of validScripts) {
            // Only generate for the currently active provider
            if (script.provider === voiceProvider) {
                 await handleGenerateSingleVoice(script.id, true);
            }
        }
    };

    const handleGenerateSfx = async () => {
        const sfxPrompts = soundEffectPrompt.split('\n').map(p => p.trim()).filter(p => p);
        if (sfxPrompts.length === 0) return;
    
        if (!elevenLabsApiKey) {
            setError('ElevenLabs API Key must be set to generate sound effects.');
            return;
        }
    
        const newJobs: SoundEffectJob[] = sfxPrompts.map(prompt => ({
            id: `sfx-${Date.now()}-${Math.random()}`,
            text: prompt,
            isGenerating: true,
        }));
    
        setSoundEffectJobs(prev => [...newJobs, ...prev]);
    
        for (const job of newJobs) {
            try {
                const audioBlob = await elevenLabsService.generateSoundEffect(elevenLabsApiKey, job.text, soundEffectDuration);
                if (audioBlob.size > 0) {
                    const audioUrl = URL.createObjectURL(audioBlob);
                    setSoundEffectJobs(prev => prev.map(j => j.id === job.id ? { ...j, audioUrl, isGenerating: false } : j));
                } else {
                    throw new Error("SFX generation returned an empty file.");
                }
            } catch (e) {
                const error = e instanceof Error ? e.message : 'An unknown error occurred.';
                setSoundEffectJobs(prev => prev.map(j => j.id === job.id ? { ...j, error, isGenerating: false } : j));
            }
        }
    
        setSoundEffectPrompt('');
    };

     const handleDownloadAllFinalAudio = async () => {
        if (finalAudioFiles.length === 0) return;

        const zip = new JSZip();
        
        const promises = finalAudioFiles.map(async (file, index) => {
            const response = await fetch(file.audioUrl!);
            const blob = await response.blob();
            const voiceName = (voices.find(v => v.id === file.voiceId)?.name || 'Unknown').replace(/ /g, '_');
            const sanitizedText = file.text.substring(0, 20).replace(/[^a-zA-Z0-9]/g, '_').toLowerCase();
            const fileName = `${index + 1}_${voiceName}_${sanitizedText}.mp3`;
            zip.file(fileName, blob);
        });

        await Promise.all(promises);

        zip.generateAsync({ type: "blob" }).then(content => {
            const link = document.createElement("a");
            link.href = URL.createObjectURL(content);
            link.download = `ai_voices_${Date.now()}.zip`;
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
        });
    };
    
    const handleSetVoiceApiKey = () => {
        if (voiceProvider === 'elevenlabs') {
            const key = tempElevenLabsKeyInput.trim();
            setElevenLabsApiKey(key);
            localStorage.setItem('elevenLabsApiKey', key);
        } else {
            const key = tempWellSaidKeyInput.trim();
            setWellSaidApiKey(key);
            localStorage.setItem('wellSaidApiKey', key);
        }
    };

    // --- AI Script Tuner Functions ---
    const handleTuneScript = async () => {
        if (!scriptTunerInput.trim()) return;

        setIsTuningScript(true);
        setScriptTunerOutput('');
        setError(null);
        
        const tone = selectedTone === 'Custom' ? customTone : selectedTone;
        if (!tone.trim()) {
            setError("Please select a tone or provide a custom tone description.");
            setIsTuningScript(false);
            return;
        }

        try {
            const tunedScript = await geminiService.perfectScriptForAI(scriptTunerInput, tone);
            setScriptTunerOutput(tunedScript);
        } catch (e) {
            setError(e instanceof Error ? e.message : 'An unknown error occurred while tuning the script.');
        } finally {
            setIsTuningScript(false);
        }
    };
    
    const handleCopyToClipboard = (text: string) => {
        if (!text || isCopied) return;
        navigator.clipboard.writeText(text).then(() => {
            setIsCopied(true);
            setTimeout(() => setIsCopied(false), 2000);
        });
    };

    // --- Bulk Script Writer Functions ---
    const handleUpdateBulkJob = (id: string, updates: Partial<BulkScriptJob>) => {
        setBulkScriptJobs(prev => prev.map(job => job.id === id ? { ...job, ...updates } : job));
    };
    
    const handleBulkScriptApiProviderChange = (provider: BulkScriptApiProvider) => {
        setBulkScriptApiProvider(provider);
        localStorage.setItem('bulkScriptApiProvider', provider);
    }

    const handleSaveBulkScriptKey = () => {
        const key = tempBulkScriptApiKey.trim();
        const provider = bulkScriptApiProvider;
        if (key) {
            localStorage.setItem(`bulk_script_api_key_${provider}`, key);
            setCustomScriptApiKeys(prev => ({...prev, [provider]: key }));
            alert('API Key saved!');
            handleFetchAndSetModels(provider, key);
        } else {
            localStorage.removeItem(`bulk_script_api_key_${provider}`);
            setCustomScriptApiKeys(prev => {
                const newKeys = {...prev};
                delete newKeys[provider as keyof typeof newKeys];
                return newKeys;
            });
            setBulkScriptModels([]);
            setBulkScriptSelectedModel('');
            setBulkScriptModelError(null);
            alert('API Key cleared!');
        }
    }
    
    const handleBulkScriptModelChange = (e: React.ChangeEvent<HTMLSelectElement>) => {
        const model = e.target.value;
        setBulkScriptSelectedModel(model);
        localStorage.setItem(`bulk_script_model_${bulkScriptApiProvider}`, model);
    };

    const handleAnalyzeTitles = async () => {
        const titles = bulkScriptTitles.split('\n').map(t => t.trim()).filter(Boolean);
        if (titles.length === 0) {
            setError("Please provide at least one title.");
            return;
        }
        
        const provider = bulkScriptApiProvider;
        const apiKey = customScriptApiKeys[provider as keyof typeof customScriptApiKeys];

        if (provider === 'gemini' && !apiKey) {
            if (!spendCredits(titles.length)) return;
        } else if (provider !== 'gemini' && (!apiKey || !bulkScriptSelectedModel)) {
            const providerName = BULK_SCRIPT_PROVIDERS.find(p => p.id === provider)?.name || 'the selected provider';
            setError(`An API key and a selected model for ${providerName} must be set to continue.`);
            return;
        }
    
        setError(null);
        setBulkScriptJobs([]); // Clear previous final scripts
    
        const initialJobs: BulkScriptJob[] = titles.map(title => ({
            id: `bscript-${Date.now()}-${Math.random()}`,
            title,
            status: 'analyzing',
            wordCount: bulkScriptWordCount,
        }));
    
        setBulkScriptJobs(initialJobs);
        setBulkScriptStep('review');
    
        for (const job of initialJobs) {
            try {
                let outlineData;
                if(provider === 'gemini'){
                    outlineData = await geminiService.generateOutlineForTitle(job.title, job.wordCount, apiKey);
                } else {
                    outlineData = await multiProviderService.generateOutlineForTitle(job.title, job.wordCount, apiKey!, provider, bulkScriptSelectedModel);
                }
                handleUpdateBulkJob(job.id, { ...outlineData, status: 'review' });
            } catch (e) {
                const error = e instanceof Error ? e.message : 'Failed to generate outline.';
                handleUpdateBulkJob(job.id, { error, status: 'error' });
            }
        }
    };
    
    const handleGenerateBulkScripts = async () => {
        const provider = bulkScriptApiProvider;
        const apiKey = customScriptApiKeys[provider as keyof typeof customScriptApiKeys];
        const providerName = BULK_SCRIPT_PROVIDERS.find(p => p.id === provider)?.name || 'the selected provider';

        let jobsToProcess: BulkScriptJob[] = [];

        if (bulkScriptMode === 'manual') {
            const titles = bulkScriptTitles.split('\n').map(t => t.trim()).filter(Boolean);
            if (titles.length === 0 || !bulkScriptGuides.trim() || !bulkScriptOutline.trim()) {
                setError("Please provide titles, guides, and an outline to generate scripts.");
                return;
            }
            
            if(provider === 'gemini' && !apiKey) {
                if (!spendCredits(titles.length)) return;
            } else if (provider !== 'gemini' && (!apiKey || !bulkScriptSelectedModel)) {
                setError(`An API key and a selected model for ${providerName} must be set to continue.`);
                return;
            }

            jobsToProcess = titles.map(title => ({
                id: `bscript-${Date.now()}-${Math.random()}`,
                title,
                guides: bulkScriptGuides,
                outline: bulkScriptOutline,
                wordCount: 0, // Set to 0 to signal inference for manual mode
                status: 'pending',
            }));
            setBulkScriptJobs(jobsToProcess);
        } else { // auto-outline mode
            const jobsInReview = bulkScriptJobs.filter(j => j.status === 'review' || j.status === 'complete'); // allow re-generation
             if (jobsInReview.length === 0) {
                setError("No outlines have been approved for script generation.");
                return;
            }
            if(provider === 'gemini' && !apiKey) {
                if (!spendCredits(jobsInReview.length)) return;
            } else if (provider !== 'gemini' && (!apiKey || !bulkScriptSelectedModel)) {
                setError(`An API key and a selected model for ${providerName} must be set to continue.`);
                return;
            }

            jobsToProcess = jobsInReview;
            setBulkScriptStep('input'); // Return to the main view to see results
        }

        setError(null);

        for (const job of jobsToProcess) {
            handleUpdateBulkJob(job.id, { status: 'generating' });
            try {
                let script;
                if(provider === 'gemini') {
                    script = await geminiService.generateSingleScript(job.title, job.guides!, job.outline!, job.wordCount, apiKey);
                } else {
                    script = await multiProviderService.generateSingleScript(job.title, job.guides!, job.outline!, job.wordCount, apiKey!, provider, bulkScriptSelectedModel);
                }
                
                handleUpdateBulkJob(job.id, { script, status: 'complete' });
            } catch (e) {
                const error = e instanceof Error ? e.message : 'An unknown error occurred.';
                handleUpdateBulkJob(job.id, { error, status: 'error' });
            }
        }
    };

    const handleDownloadScriptAsTxt = (scriptText: string, title: string) => {
        const sanitizedTitle = title.replace(/[^a-z0-9]/gi, '_').toLowerCase();
        const blob = new Blob([scriptText], { type: 'text/plain;charset=utf-8' });
        saveAs(blob, `${sanitizedTitle}.txt`);
    };

    const handleDownloadScriptAsDocx = (scriptText: string, title: string) => {
        const sanitizedTitle = title.replace(/[^a-z0-9]/gi, '_').toLowerCase();
        const doc = new docx.Document({
            sections: [{
                children: scriptText.split('\n').map(line => new docx.Paragraph({ text: line })),
            }],
        });

        docx.Packer.toBlob(doc).then(blob => {
            saveAs(blob, `${sanitizedTitle}.docx`);
        });
    };

    const handleDownloadAllScripts = async (format: 'txt' | 'docx') => {
        const zip = new JSZip();
        const successfulJobs = bulkScriptJobs.filter(job => job.script && job.status === 'complete');

        if (successfulJobs.length === 0) return;

        if (format === 'docx') {
            const docxPromises = successfulJobs.map(job => {
                const doc = new docx.Document({
                    sections: [{
                        children: (job.script || '').split('\n').map(line => new docx.Paragraph({ text: line })),
                    }],
                });
                return docx.Packer.toBlob(doc).then(blob => ({
                    blob,
                    title: job.title,
                }));
            });

            const docxBlobs = await Promise.all(docxPromises);
            docxBlobs.forEach(({ blob, title }) => {
                const sanitizedTitle = title.replace(/[^a-z0-9]/gi, '_').toLowerCase();
                zip.file(`${sanitizedTitle}.docx`, blob);
            });

        } else { // txt format
             successfulJobs.forEach(job => {
                const sanitizedTitle = job.title.replace(/[^a-z0-9]/gi, '_').toLowerCase();
                zip.file(`${sanitizedTitle}.txt`, job.script || '');
            });
        }
        
        zip.generateAsync({ type: "blob" }).then(content => {
            saveAs(content, `ai_scripts_${Date.now()}.zip`);
        });
    };

    const renderConfigOptions = (isSimple: boolean = false) => (
        <div className="space-y-4">
            <div className={`grid ${isSimple ? 'grid-cols-1' : 'grid-cols-2'} gap-4`}>
                <div>
                    <label className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Aspect Ratio</label>
                    <div className="grid grid-cols-5 gap-2">
                        {ASPECT_RATIOS.map(r => (
                            <button key={r} onClick={() => setRatio(r)} className={`p-2 rounded-lg flex flex-col justify-center items-center transition-colors ${ratio === r ? 'bg-gradient-to-br from-cyan-500 to-blue-600 text-white ring-2 ring-cyan-400' : 'bg-slate-100 dark:bg-[#0b1120] hover:bg-slate-200 dark:hover:bg-cyan-400/10 text-gray-700 dark:text-gray-300'}`} title={r}>
                                <AspectRatioIcon ratio={r} />
                                <span className="text-xs font-medium mt-1">{r}</span>
                            </button>
                        ))}
                    </div>
                </div>
                <div>
                  <label className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Image Style</label>
                  <select
                    value={imageStyle}
                    onChange={(e) => setImageStyle(e.target.value)}
                    disabled={imageStyles.length === 0}
                    className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none transition-all duration-300 disabled:opacity-50"
                  >
                      {imageStyles.length === 0 && <option>Unavailable (Add API Key)</option>}
                      {imageStyles.map(style => (
                        <option key={style} value={style}>{style}</option>
                      ))}
                  </select>
                </div>
               {!isSimple && (
                <>
                    <div>
                        <label className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Image API Provider</label>
                         <select
                            value={apiProvider}
                            onChange={(e) => setApiProvider(e.target.value as ApiProvider)}
                            className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none transition-all duration-300"
                        >
                            {IMAGE_API_PROVIDERS.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
                        </select>
                    </div>
                    <div>
                        {apiProvider !== 'gemini' && (
                            <div className="mb-4">
                                <label htmlFor="image-api-key" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">API Key</label>
                                <div className="flex items-center gap-2">
                                    <input
                                        id="image-api-key"
                                        type="password"
                                        value={tempImageApiKey}
                                        onChange={e => setTempImageApiKey(e.target.value)}
                                        placeholder={`Enter API Key`}
                                        className="flex-grow bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none"
                                    />
                                    <button
                                        onClick={handleSaveImageApiKey}
                                        className="px-4 py-2.5 rounded-lg font-semibold transition-colors bg-slate-200 dark:bg-[#252f41] text-gray-800 dark:text-white hover:bg-slate-300 dark:hover:bg-cyan-500/10"
                                    >
                                        Save
                                    </button>
                                </div>
                            </div>
                        )}
                        {apiProvider === 'together' && (
                            <div>
                                <label htmlFor="steps" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Steps ({steps})</label>
                                <input
                                    type="range"
                                    id="steps"
                                    min="1"
                                    max="50"
                                    value={steps}
                                    onChange={e => setSteps(Number(e.target.value))}
                                    className="w-full h-2 bg-slate-200 dark:bg-[#252f41] rounded-lg appearance-none cursor-pointer accent-cyan-400"
                                />
                            </div>
                        )}
                    </div>
                </>
               )}
            </div>
            {apiProvider === 'gemini' && (
                <div className="animate-fade-in pt-4 border-t border-slate-200 dark:border-[#252f41]">
                    <label className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Style Reference (Optional)</label>
                    {styleReferenceImage ? (
                        <div className="flex items-center gap-3 p-2 bg-slate-100 dark:bg-[#0b1120] rounded-lg border border-slate-200 dark:border-[#252f41]">
                            <img src={styleReferenceImage.url} alt="Style Reference" className="w-12 h-12 object-cover rounded-md" />
                            <p className="flex-1 text-sm text-gray-700 dark:text-gray-300 truncate">Using style from uploaded image.</p>
                            <button onClick={clearStyleReference} className="p-1.5 rounded-full hover:bg-red-500/10 text-gray-500 hover:text-red-500 transition-colors">
                                <CloseIcon className="h-5 w-5" />
                            </button>
                        </div>
                    ) : (
                        <button
                            onClick={() => styleReferenceFileInputRef.current?.click()}
                            className="w-full flex items-center justify-center gap-2 text-sm font-semibold bg-slate-200/80 dark:bg-[#0b1120] hover:bg-slate-200 dark:hover:bg-cyan-500/10 text-gray-700 dark:text-gray-300 py-2.5 px-3 rounded-lg border-2 border-dashed border-slate-300 dark:border-[#252f41] transition-colors"
                        >
                            <UploadIcon /> Add Image Reference
                        </button>
                    )}
                    <input type="file" ref={styleReferenceFileInputRef} onChange={handleStyleReferenceChange} accept="image/*" className="hidden" />
                </div>
            )}
        </div>
    );
    
    const renderCaptionToggle = (label: string, enabled: boolean, onChange: (value: boolean) => void) => (
         <div
            role="switch"
            aria-checked={enabled}
            onClick={() => onChange(!enabled)}
            className="flex items-center justify-between p-2 rounded-lg transition-colors cursor-pointer bg-slate-200/50 dark:bg-[#0b1120]"
        >
            <span className="text-sm font-medium text-gray-700 dark:text-gray-300">{label}</span>
            <div className={`w-10 h-5 flex items-center rounded-full p-0.5 transition-colors ${enabled ? 'bg-cyan-500' : 'bg-gray-400 dark:bg-gray-700'}`}>
                <div className={`w-4 h-4 bg-white rounded-full shadow-md transform transition-transform ${enabled ? 'translate-x-5' : 'translate-x-0'}`}/>
            </div>
        </div>
    );

    const renderVideoSettings = () => (
        <div className="bg-slate-50 dark:bg-[#172134] border border-slate-200 dark:border-[#252f41] rounded-2xl p-4 md:p-6 mt-6 animate-fade-in">
            <div className="flex justify-between items-center mb-6">
                 <div className="flex items-center gap-3">
                    <SettingsIcon/>
                    <h3 className="text-xl font-bold text-gray-900 dark:text-white">Video & Caption Settings</h3>
                </div>
                <div className="flex items-center gap-2">
                    <button onClick={undoVideoConfig} disabled={videoConfigState.past.length === 0} className="p-2 rounded-full bg-slate-100 dark:bg-[#0b1120] hover:bg-slate-200 dark:hover:bg-cyan-500/10 disabled:opacity-50 disabled:cursor-not-allowed transition-colors" aria-label="Undo">
                        <UndoIcon className="h-5 w-5 text-gray-800 dark:text-white" />
                    </button>
                    <button onClick={redoVideoConfig} disabled={videoConfigState.future.length === 0} className="p-2 rounded-full bg-slate-100 dark:bg-[#0b1120] hover:bg-slate-200 dark:hover:bg-cyan-500/10 disabled:opacity-50 disabled:cursor-not-allowed transition-colors" aria-label="Redo">
                        <RedoIcon className="h-5 w-5 text-gray-800 dark:text-white" />
                    </button>
                </div>
            </div>
           
            <div className="space-y-6">
                {/* --- Section 1: General --- */}
                <div>
                    <h4 className="text-lg font-semibold text-gray-800 dark:text-white mb-3 border-b border-slate-200 dark:border-[#252f41] pb-2">General</h4>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div>
                            <label className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">AI Voiceover</label>
                            <VoiceSelector
                                voices={voices}
                                selectedValue={videoConfig.voice}
                                onSelect={(voiceId) => setVideoConfig(c => ({...c, voice: voiceId}))}
                                onPreview={handlePreviewVoice}
                                previewStatus={previewStatus}
                                isFetching={isFetchingVoices}
                                isUsingApiKey={isUsingVoiceApiKey}
                                apiKeyProviderName={voiceProvider === 'elevenlabs' ? 'ElevenLabs' : 'WellSaid Labs'}
                                onOpenApiKeyModal={() => setIsApiKeyModalOpen(true)}
                                uiVariant={theme}
                            />
                        </div>
                        <div>
                            <label className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Scene Transition</label>
                             <select value={videoConfig.sceneTransition} onChange={(e) => setVideoConfig(c => ({...c, sceneTransition: e.target.value as SceneTransition}))} className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-base focus:ring-2 focus:ring-cyan-500 outline-none">
                                {SCENE_TRANSITIONS.map(t => <option key={t} value={t}>{t.split('-').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ')}</option>)}
                            </select>
                        </div>
                    </div>
                </div>
    
                {/* --- Section 2: Caption Layout --- */}
                <div>
                     <h4 className="text-lg font-semibold text-gray-800 dark:text-white mb-3 border-b border-slate-200 dark:border-[#252f41] pb-2">Caption Layout & Animation</h4>
                     <div
                        role="switch"
                        aria-checked={videoConfig.captionsEnabled}
                        onClick={() => setVideoConfig(c => ({...c, captionsEnabled: !c.captionsEnabled}))}
                        className="flex items-center justify-between p-3 rounded-lg transition-colors cursor-pointer bg-slate-100 dark:bg-[#0b1120] mb-4"
                    >
                        <span className="font-semibold text-gray-700 dark:text-gray-200">Enable Captions</span>
                        <div className={`w-11 h-6 flex items-center rounded-full p-1 transition-colors ${videoConfig.captionsEnabled ? 'bg-cyan-500' : 'bg-gray-400 dark:bg-gray-700'}`}>
                            <div className={`w-4 h-4 bg-white rounded-full shadow-md transform transition-transform ${videoConfig.captionsEnabled ? 'translate-x-5' : 'translate-x-0'}`}/>
                        </div>
                    </div>
                    {videoConfig.captionsEnabled && (
                        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 animate-fade-in">
                            <div>
                                <label className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Animation</label>
                                 <select value={videoConfig.captionAnimation} onChange={(e) => setVideoConfig(c => ({...c, captionAnimation: e.target.value as CaptionAnimation}))} className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-base focus:ring-2 focus:ring-cyan-500 outline-none">
                                    {CAPTION_ANIMATIONS.map(a => <option key={a} value={a}>{a.split('-').map(w => w.charAt(0).toUpperCase() + w.slice(1)).join(' ')}</option>)}
                                </select>
                            </div>
                            <div>
                                <label htmlFor="wordsPerLine" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Words Per Line ({videoConfig.captionWordsPerLine})</label>
                                <input id="wordsPerLine" type="range" min="3" max="15" value={videoConfig.captionWordsPerLine} onChange={(e) => setVideoConfig(c => ({...c, captionWordsPerLine: Number(e.target.value)}))} className="w-full h-2 bg-slate-300 dark:bg-[#252f41] rounded-lg appearance-none cursor-pointer accent-cyan-400" />
                            </div>
                        </div>
                    )}
                </div>
    
                {/* --- Section 3: Caption Style --- */}
                {videoConfig.captionsEnabled && (
                <div className="animate-fade-in">
                    <h4 className="text-lg font-semibold text-gray-800 dark:text-white mb-4 border-b border-slate-200 dark:border-[#252f41] pb-2">Caption Style</h4>
                    
                    <div className="mb-6">
                        <label className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Caption Font</label>
                        <select 
                            value={videoConfig.captionFont} 
                            onChange={(e) => setVideoConfig(c => ({...c, captionFont: e.target.value}))} 
                            className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-base focus:ring-2 focus:ring-cyan-500 outline-none"
                            style={{ fontFamily: videoConfig.captionFont }}
                        >
                            {CAPTION_FONTS.map(font => <option key={font} value={font} style={{fontFamily: font}}>{font}</option>)}
                        </select>
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-6">
                        {renderCaptionToggle('Background', videoConfig.backgroundEnabled, (val) => setVideoConfig(c => ({...c, backgroundEnabled: val})))}
                        {renderCaptionToggle('Stroke', videoConfig.strokeEnabled, (val) => setVideoConfig(c => ({...c, strokeEnabled: val})))}
                        {renderCaptionToggle('Shadow', videoConfig.shadowEnabled, (val) => setVideoConfig(c => ({...c, shadowEnabled: val})))}
                    </div>
                    
                    <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 mb-6">
                        <div className="flex flex-col items-start gap-2">
                            <label htmlFor="captionColor" className="text-sm font-medium text-gray-600 dark:text-gray-400">Text Color</label>
                            <input id="captionColor" type="color" value={videoConfig.captionColor} onChange={(e) => setVideoConfig(c => ({...c, captionColor: e.target.value}))} className="w-full h-10 bg-transparent dark:bg-[#0b1120] border-2 border-slate-200 dark:border-[#252f41] rounded-lg cursor-pointer p-1" />
                        </div>
                         {videoConfig.backgroundEnabled && (
                            <div className="flex flex-col items-start gap-2 animate-fade-in">
                                <label htmlFor="captionBgColor" className="text-sm font-medium text-gray-600 dark:text-gray-400">BG Color</label>
                                <input id="captionBgColor" type="color" value={videoConfig.captionBackgroundColor} onChange={(e) => setVideoConfig(c => ({...c, captionBackgroundColor: e.target.value}))} className="w-full h-10 bg-transparent dark:bg-[#0b1120] border-2 border-slate-200 dark:border-[#252f41] rounded-lg cursor-pointer p-1" />
                            </div>
                        )}
                        {videoConfig.strokeEnabled && (
                            <div className="flex flex-col items-start gap-2 animate-fade-in">
                                <label htmlFor="captionStrokeColor" className="text-sm font-medium text-gray-600 dark:text-gray-400">Stroke Color</label>
                                <input id="captionStrokeColor" type="color" value={videoConfig.captionStrokeColor} onChange={(e) => setVideoConfig(c => ({...c, captionStrokeColor: e.target.value}))} className="w-full h-10 bg-transparent dark:bg-[#0b1120] border-2 border-slate-200 dark:border-[#252f41] rounded-lg cursor-pointer p-1" />
                            </div>
                        )}
                        {videoConfig.shadowEnabled && (
                            <div className="flex flex-col items-start gap-2 animate-fade-in">
                                <label htmlFor="captionShadowColor" className="text-sm font-medium text-gray-600 dark:text-gray-400">Shadow Color</label>
                                <input id="captionShadowColor" type="color" value={videoConfig.captionShadowColor} onChange={(e) => setVideoConfig(c => ({...c, captionShadowColor: e.target.value}))} className="w-full h-10 bg-transparent dark:bg-[#0b1120] border-2 border-slate-200 dark:border-[#252f41] rounded-lg cursor-pointer p-1" />
                            </div>
                        )}
                    </div>
                    
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                        {videoConfig.backgroundEnabled && (
                            <div className="animate-fade-in">
                                <label htmlFor="bgOpacity" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Background Opacity ({Math.round(videoConfig.captionBackgroundOpacity * 100)}%)</label>
                                <input id="bgOpacity" type="range" min="0" max="1" step="0.05" value={videoConfig.captionBackgroundOpacity} onChange={(e) => setVideoConfig(c => ({...c, captionBackgroundOpacity: Number(e.target.value)}))} className="w-full h-2 bg-slate-300 dark:bg-[#252f41] rounded-lg appearance-none cursor-pointer accent-cyan-400" />
                            </div>
                        )}
                        {videoConfig.strokeEnabled && (
                            <div className="animate-fade-in">
                                <label htmlFor="strokeWidth" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Stroke Width ({videoConfig.captionStrokeWidth}px)</label>
                                <input id="strokeWidth" type="range" min="0" max="10" step="0.5" value={videoConfig.captionStrokeWidth} onChange={(e) => setVideoConfig(c => ({...c, captionStrokeWidth: Number(e.target.value)}))} className="w-full h-2 bg-slate-300 dark:bg-[#252f41] rounded-lg appearance-none cursor-pointer accent-cyan-400" />
                            </div>
                        )}
                        {videoConfig.shadowEnabled && (
                            <div className="md:col-span-2 animate-fade-in">
                                 <label htmlFor="shadowBlur" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Shadow Blur ({videoConfig.captionShadowBlur}px)</label>
                                <input id="shadowBlur" type="range" min="0" max="25" value={videoConfig.captionShadowBlur} onChange={(e) => setVideoConfig(c => ({...c, captionShadowBlur: Number(e.target.value)}))} className="w-full h-2 bg-slate-300 dark:bg-[#252f41] rounded-lg appearance-none cursor-pointer accent-cyan-400" />
                            </div>
                        )}
                    </div>
                </div>
                )}
            </div>
    
            <button onClick={handleGenerateVideo} disabled={isGeneratingVideo || activeResults.filter(r => r.imageUrl).length === 0} className="mt-8 w-full flex items-center justify-center gap-3 bg-green-600 hover:bg-green-700 text-white font-bold py-3 px-5 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-all">
                {isGeneratingVideo ? (
                     <>
                        <div role="status" className="inline-block h-5 w-5 animate-spin rounded-full border-2 border-solid border-current border-r-transparent align-[-0.125em] motion-reduce:animate-[spin_1.5s_linear_infinite]" />
                        Rendering Video...
                    </>
                ) : (
                    <>
                        <VideoIcon/> Generate AI Video
                    </>
                )}
            </button>
        </div>
    );
    
    const renderVoiceGeneratorResultsPanel = () => (
        <div className="bg-slate-50 dark:bg-[#172134] p-6 rounded-2xl border border-slate-200 dark:border-[#252f41] animate-fade-in h-full max-h-[calc(100vh-10rem)]">
            <div className="flex justify-between items-center mb-4">
                <h3 className="text-xl font-bold text-gray-900 dark:text-white">Final Audio Results</h3>
                <button
                    onClick={handleDownloadAllFinalAudio}
                    disabled={finalAudioFiles.length === 0}
                    className="flex items-center gap-1.5 text-sm font-semibold bg-slate-200 dark:bg-[#0b1120] hover:bg-slate-300/80 dark:hover:bg-cyan-500/10 py-1.5 px-3 rounded-lg transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                >
                    <ZipIcon /> Download All
                </button>
            </div>
            <div className="space-y-3 overflow-y-auto max-h-[calc(100%-4rem)] pr-2">
                {finalAudioFiles.map((file) => {
                    const voiceName = voices.find(v => v.id === file.voiceId)?.name || 'Unknown Voice';
                    const sanitizedText = file.text.substring(0, 20).replace(/[^a-zA-Z0-9]/g, '_').toLowerCase();
                    const downloadFilename = `${voiceName.replace(/ /g, '_')}_${sanitizedText}.mp3`;
                    return (
                        <div key={file.id} className="p-3 bg-slate-100 dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg">
                            <div className="flex justify-between items-start gap-2">
                                <div>
                                    <p className="text-xs text-gray-500 dark:text-gray-400 font-semibold uppercase">{voiceName}</p>
                                    <p className="text-sm text-gray-700 dark:text-gray-300 line-clamp-2 my-1">{file.text}</p>
                                </div>
                                <a href={file.audioUrl} download={downloadFilename} title="Download MP3" className="flex-shrink-0 p-2.5 bg-slate-200 dark:bg-slate-800/50 text-gray-600 dark:text-cyan-300 hover:bg-slate-300 dark:hover:bg-cyan-500/30 rounded-md transition-colors">
                                    <DownloadIcon />
                                </a>
                            </div>
                            <audio controls src={file.audioUrl} className="w-full h-9 mt-2"></audio>
                        </div>
                    );
                })}
            </div>
        </div>
    );


    const renderVoiceGenerator = () => {
        const anyScriptGenerating = voiceScripts.some(s => s.isGenerating);
        const validScriptsCount = voiceScripts.filter(s => s.text.trim() && s.voiceId && s.provider === voiceProvider).length;
        const sfxPromptsCount = soundEffectPrompt.split('\n').map(p => p.trim()).filter(p => p).length;


        return (
            <div className="space-y-4">
                <p className="text-sm text-gray-600 dark:text-gray-400">
                    Create audio from multiple scripts with different voices.
                </p>

                <div className="flex space-x-1 bg-slate-100 dark:bg-[#0b1120] p-1 rounded-xl">
                    <button onClick={() => setVoiceGeneratorMode('voice')} className={`w-full flex items-center justify-center gap-2 text-sm font-semibold py-1.5 px-2 rounded-lg transition-colors ${voiceGeneratorMode === 'voice' ? 'gradient-button text-white' : 'text-gray-600 dark:text-gray-300 hover:bg-slate-200 dark:hover:bg-cyan-500/10'}`}>
                        <MicrophoneIcon className="h-5 w-5" /> Voices
                    </button>
                    <button onClick={() => setVoiceGeneratorMode('sfx')} className={`w-full flex items-center justify-center gap-2 text-sm font-semibold py-1.5 px-2 rounded-lg transition-colors ${voiceGeneratorMode === 'sfx' ? 'gradient-button text-white' : 'text-gray-600 dark:text-gray-300 hover:bg-slate-200 dark:hover:bg-cyan-500/10'}`}>
                        <SoundWaveIcon className="h-5 w-5" /> Sound FX
                    </button>
                </div>
                
                {error && (
                    <div className="bg-red-100 dark:bg-red-900/30 border border-red-500/50 text-red-700 dark:text-red-300 px-4 py-3 rounded-lg relative" role="alert">
                        <strong className="font-bold">Error: </strong>
                        <span className="block sm:inline">{error}</span>
                         <button onClick={() => setError(null)} className="absolute top-0 bottom-0 right-0 px-4 py-3">
                            <CloseIcon className="h-5 w-5"/>
                        </button>
                    </div>
                )}
                
                <div className="space-y-4 pt-4 border-t border-slate-200 dark:border-[#252f41]">
                    <h3 className="text-base font-semibold text-gray-700 dark:text-white">Configuration</h3>
                    {voiceGeneratorMode === 'voice' && (
                        <div className="flex space-x-1 bg-slate-200 dark:bg-[#0b1120] p-1 rounded-xl mb-4">
                            <button onClick={() => setVoiceProvider('elevenlabs')} className={`w-full text-sm font-semibold py-1.5 px-2 rounded-lg transition-colors ${voiceProvider === 'elevenlabs' ? 'bg-white dark:bg-[#172134] text-gray-800 dark:text-white shadow-sm' : 'text-gray-600 dark:text-gray-400 hover:bg-slate-300/50 dark:hover:bg-cyan-500/10'}`}>
                                ElevenLabs
                            </button>
                            <button onClick={() => setVoiceProvider('wellsaid')} className={`w-full text-sm font-semibold py-1.5 px-2 rounded-lg transition-colors ${voiceProvider === 'wellsaid' ? 'bg-white dark:bg-[#172134] text-gray-800 dark:text-white shadow-sm' : 'text-gray-600 dark:text-gray-400 hover:bg-slate-300/50 dark:hover:bg-cyan-500/10'}`}>
                                WellSaid Labs
                            </button>
                        </div>
                    )}
                    <div className="flex items-center gap-2">
                        {voiceProvider === 'elevenlabs' && (
                            <input
                                type="password"
                                value={tempElevenLabsKeyInput}
                                onChange={e => setTempElevenLabsKeyInput(e.target.value)}
                                placeholder="ElevenLabs API Key"
                                className="flex-grow bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none"
                            />
                        )}
                         {voiceProvider === 'wellsaid' && voiceGeneratorMode === 'voice' && (
                            <input
                                type="password"
                                value={tempWellSaidKeyInput}
                                onChange={e => setTempWellSaidKeyInput(e.target.value)}
                                placeholder="WellSaid Labs API Key"
                                className="flex-grow bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none"
                            />
                        )}
                         {voiceGeneratorMode === 'sfx' && (
                            <input
                                type="password"
                                value={tempElevenLabsKeyInput}
                                onChange={e => setTempElevenLabsKeyInput(e.target.value)}
                                placeholder="ElevenLabs API Key (for SFX)"
                                className="flex-grow bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none"
                            />
                         )}
                        <button
                            onClick={handleSetVoiceApiKey}
                            className={`flex items-center justify-center px-3 py-2 rounded-lg font-semibold transition-colors ${isUsingVoiceApiKey ? 'bg-green-500/20 text-green-400' : 'bg-slate-200 dark:bg-[#252f41] hover:bg-slate-300 dark:hover:bg-cyan-500/10 text-gray-800 dark:text-white'}`}
                        >
                            {isUsingVoiceApiKey ? <CheckCircleIcon className="h-5 w-5" /> : 'Set'}
                        </button>
                    </div>
                </div>

                <div className="flex flex-col gap-6">
                    <div className="flex-1 min-w-0">
                        {voiceGeneratorMode === 'voice' && (
                        <div className="space-y-4 animate-fade-in">
                            <div className="space-y-4 pt-4 border-t border-slate-200 dark:border-[#252f41]">
                                <h3 className="text-base font-semibold text-gray-700 dark:text-white">Scripts & Voices</h3>
                                <div className="space-y-4 max-h-[40vh] overflow-y-auto pr-2">
                                    {voiceScripts.map((script, index) => {
                                        const wordCount = script.text.trim().split(/\s+/).filter(Boolean).length;
                                        const characterCount = script.text.length;
                                        return (
                                        <div key={script.id} className="p-4 bg-white dark:bg-[#0b1120] rounded-xl border border-slate-200 dark:border-[#252f41] relative shadow-md">
                                            <div className="flex justify-between items-center mb-2">
                                                <h4 className="font-semibold text-gray-800 dark:text-white">Script {index + 1}</h4>
                                                <button onClick={() => handleRemoveVoiceScript(script.id)} className="p-1.5 text-gray-500 dark:text-gray-400 hover:text-red-500 hover:bg-red-500/10 rounded-full transition-colors">
                                                    <TrashIcon />
                                                </button>
                                            </div>
                                            <div className="space-y-3">
                                                <div>
                                                    <label className="block text-xs font-medium text-gray-600 dark:text-gray-400 mb-1">Voice Character</label>
                                                    <VoiceSelector 
                                                        voices={voices}
                                                        selectedValue={script.voiceId}
                                                        onSelect={voiceId => handleUpdateVoiceScript(script.id, { voiceId, provider: voiceProvider })}
                                                        onPreview={handlePreviewVoice}
                                                        previewStatus={previewStatus}
                                                        isFetching={isFetchingVoices}
                                                        isUsingApiKey={isUsingVoiceApiKey}
                                                        apiKeyProviderName={voiceProvider === 'elevenlabs' ? 'ElevenLabs' : 'WellSaid Labs'}
                                                        onOpenApiKeyModal={() => {}}
                                                        uiVariant={theme}
                                                    />
                                                </div>
                                                {voiceProvider === 'elevenlabs' && (
                                                <div>
                                                    <label className="block text-xs font-medium text-gray-600 dark:text-gray-400 mb-1">Speed</label>
                                                    <div className="flex items-center gap-2">
                                                        <span className="text-xs text-gray-500 dark:text-gray-400">Slower</span>
                                                        <input
                                                            type="range"
                                                            min="0" max="100"
                                                            value={script.speed}
                                                            onChange={e => handleUpdateVoiceScript(script.id, { speed: Number(e.target.value) })}
                                                            className="w-full h-2 bg-slate-200 dark:bg-[#252f41] rounded-lg appearance-none cursor-pointer accent-cyan-400"
                                                        />
                                                        <span className="text-xs text-gray-500 dark:text-gray-400">Faster</span>
                                                    </div>
                                                </div>
                                                )}
                                            </div>
                                            <div className="mt-3">
                                                <textarea
                                                    value={script.text}
                                                    onChange={e => handleUpdateVoiceScript(script.id, { text: e.target.value })}
                                                    rows={5}
                                                    placeholder="Enter script text here, or paste a tuned script with SSML tags..."
                                                    className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-sm focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none resize-y"
                                                />
                                                <div className="flex justify-between items-center mt-1 text-xs text-gray-500 dark:text-gray-400">
                                                    <button onClick={() => voiceScriptFileInputRefs.current[script.id]?.click()} className="font-medium text-cyan-600 dark:text-cyan-400 hover:underline">
                                                        Upload...
                                                    </button>
                                                    <input type="file" ref={el => { if(el) voiceScriptFileInputRefs.current[script.id] = el }} onChange={(e) => handleFileChange(e, () => {}, script.id)} accept=".txt,.docx" className="hidden"/>
                                                    <div className="flex items-center gap-3">
                                                        <span>{wordCount} Words</span>
                                                        <span>{characterCount} Chars</span>
                                                    </div>
                                                </div>
                                            </div>
                                            <div className="mt-3 space-y-3">
                                                {script.audioUrl && (
                                                    <div className={`p-1 rounded-lg transition-all duration-300 ease-in-out ${script.justUpdated ? 'bg-green-500/20 ring-2 ring-green-500/50 shadow-md' : 'bg-transparent'}`}>
                                                        <audio controls src={script.audioUrl} className="w-full h-10"></audio>
                                                    </div>
                                                )}

                                                {/* Show Generate button if no audio OR if it's been modified */}
                                                {(!script.audioUrl || script.isModified) && (
                                                    <button
                                                        onClick={() => handleGenerateSingleVoice(script.id)}
                                                        disabled={!script.text.trim() || !script.voiceId || script.isGenerating || (anyScriptGenerating && !script.isGenerating)}
                                                        className="w-full bg-cyan-600/10 text-cyan-700 dark:bg-cyan-500/10 dark:text-cyan-300 hover:bg-cyan-600/20 dark:hover:bg-cyan-500/20 font-semibold px-4 py-2 rounded-md transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                                                    >
                                                        {script.isGenerating ? 'Generating...' : (script.isModified ? 'Regenerate Audio' : 'Generate Audio')}
                                                    </button>
                                                )}
                                            </div>
                                            {script.error && <p className="text-red-500 text-sm mt-2">{script.error}</p>}
                                            {script.isGenerating && (
                                                <div className="absolute inset-0 bg-white/50 dark:bg-black/50 backdrop-blur-sm flex items-center justify-center rounded-lg">
                                                    <Spinner />
                                                </div>
                                            )}
                                        </div>
                                    )})}
                                </div>
                            </div>

                            <div className="flex flex-col gap-2 pt-4 border-t border-slate-200 dark:border-[#252f41]">
                                <button
                                    onClick={handleAddVoiceScript}
                                    className="w-full bg-slate-200 dark:bg-[#0b1120] hover:bg-slate-300/80 dark:hover:bg-cyan-500/10 text-gray-800 dark:text-white font-semibold px-4 py-2.5 rounded-lg transition-colors"
                                >
                                    + Add Another Script
                                </button>
                                <button
                                    onClick={handleGenerateAllVoices}
                                    disabled={anyScriptGenerating || validScriptsCount === 0 || !isUsingVoiceApiKey}
                                    className="w-full gradient-button text-white font-bold px-4 py-2.5 rounded-lg transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                                >
                                    Generate All ({validScriptsCount}) Valid Script(s)
                                </button>
                            </div>
                        </div>
                        )}

                        {voiceGeneratorMode === 'sfx' && (
                             <div className="space-y-4 animate-fade-in">
                                <div className="space-y-4 pt-4 border-t border-slate-200 dark:border-[#252f41]">
                                     <h3 className="text-base font-semibold text-gray-700 dark:text-white">Generate Sound Effect(s)</h3>
                                     <p className="text-sm text-gray-600 dark:text-gray-400">Describe the sounds you want to create (one per line). Powered by ElevenLabs.</p>
                                     <div className="flex flex-col gap-4">
                                         <div>
                                            <textarea
                                                 value={soundEffectPrompt}
                                                 onChange={e => setSoundEffectPrompt(e.target.value)}
                                                 rows={5}
                                                 placeholder={"a spaceship door opening\nlaser blast\nfootsteps on gravel"}
                                                 className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none resize-y"
                                                 disabled={isGeneratingSfx}
                                             />
                                             <div className="flex justify-between items-center mt-2 pr-1 text-xs text-gray-500 dark:text-gray-400">
                                                <span/> {/* Spacer */}
                                                <button onClick={() => sfxFileInputRef.current?.click()} className="font-medium text-cyan-600 dark:text-cyan-400 hover:underline">
                                                    Upload from .txt / .docx...
                                                </button>
                                                <input type="file" ref={sfxFileInputRef} onChange={(e) => handleFileChange(e, setSoundEffectPrompt)} accept=".txt,.docx" className="hidden"/>
                                             </div>
                                         </div>
                                          <div>
                                              <label htmlFor="sfx-duration" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Duration ({soundEffectDuration.toFixed(1)}s)</label>
                                              <input
                                                id="sfx-duration"
                                                type="range"
                                                min="3"
                                                max="22"
                                                step="0.5"
                                                value={soundEffectDuration}
                                                onChange={e => setSoundEffectDuration(Number(e.target.value))}
                                                className="w-full h-2 bg-slate-200 dark:bg-[#252f41] rounded-lg appearance-none cursor-pointer accent-cyan-400"
                                                disabled={isGeneratingSfx}
                                              />
                                          </div>
                                         <button
                                             onClick={handleGenerateSfx}
                                             disabled={!soundEffectPrompt.trim() || isGeneratingSfx || !elevenLabsApiKey}
                                             className="w-full gradient-button font-bold py-2.5 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed"
                                         >
                                            {isGeneratingSfx ? "Generating..." : `Generate SFX ${sfxPromptsCount > 0 ? `(${sfxPromptsCount})` : ''}`}
                                         </button>
                                     </div>
                                </div>

                                {soundEffectJobs.length > 0 && (
                                    <div className="space-y-4 pt-4 border-t border-slate-200 dark:border-[#252f41]">
                                        <h3 className="text-base font-semibold text-gray-700 dark:text-white">Generated Sound Effects</h3>
                                        <div className="space-y-3 max-h-60 overflow-y-auto pr-2">
                                            {soundEffectJobs.map(job => (
                                                <div key={job.id} className="p-3 bg-slate-100 dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg relative">
                                                    <p className="font-semibold text-gray-800 dark:text-white truncate">{job.text}</p>
                                                    {job.isGenerating && <p className="text-sm text-cyan-500 dark:text-cyan-400 mt-2">Generating audio...</p>}
                                                    {job.error && <p className="text-sm text-red-500 dark:text-red-400 mt-2">{job.error}</p>}
                                                    {job.audioUrl && (
                                                        <div className="flex items-center gap-2 mt-2">
                                                            <audio controls src={job.audioUrl} className="w-full h-9"></audio>
                                                            <a href={job.audioUrl} download={`sfx_${job.text.replace(/ /g, '_')}.mp3`} title="Download MP3" className="p-2 bg-slate-200 dark:bg-cyan-500/20 text-gray-700 dark:text-cyan-300 hover:bg-slate-300 dark:hover:bg-cyan-500/30 rounded-md transition-colors">
                                                                <DownloadIcon />
                                                            </a>
                                                        </div>
                                                    )}
                                                </div>
                                            ))}
                                        </div>
                                    </div>
                                )}
                             </div>
                        )}
                    </div>
                </div>

                <p className="text-center text-xs text-gray-500 dark:text-gray-400 pt-4">
                    Generated audio is not stored. Refreshing the page will clear results.
                </p>
            </div>
        );
    }
    
    const renderScriptTuner = () => {
        const TONE_OPTIONS = ['Documentary', 'Advertisement', 'Storyteller', 'Meditative', 'Excited', 'Horror', 'Motivational', 'Custom'];
        
        return (
            <div className="space-y-4">
                <p className="text-sm text-gray-600 dark:text-gray-400">
                    Enhance scripts for AI voice generation. This tool embeds special SSML tags for pacing, emotion, and emphasis to make AI voices sound more natural. Copy the tuned script into the AI Voice Generator to hear the result.
                </p>

                {error && (
                     <div className="bg-red-100 dark:bg-red-900/30 border border-red-500/50 text-red-700 dark:text-red-300 px-4 py-3 rounded-lg relative" role="alert">
                        <strong className="font-bold">Error: </strong>
                        <span className="block sm:inline">{error}</span>
                         <button onClick={() => setError(null)} className="absolute top-0 bottom-0 right-0 px-4 py-3">
                            <CloseIcon className="h-5 w-5"/>
                        </button>
                    </div>
                )}
                
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-4 border-t border-slate-200 dark:border-[#252f41]">
                    {/* Input Column */}
                    <div className="space-y-4">
                        <h3 className="text-lg font-semibold text-gray-800 dark:text-white">Your Original Script</h3>
                        <div>
                            <textarea
                                value={scriptTunerInput}
                                onChange={e => setScriptTunerInput(e.target.value)}
                                rows={10}
                                placeholder="Paste your script here..."
                                className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-3 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none transition-all duration-300 resize-y"
                                disabled={isTuningScript}
                            />
                             <div className="flex justify-end mt-1">
                                <button 
                                    onClick={() => scriptTunerFileInputRef.current?.click()} 
                                    className="text-sm font-medium text-cyan-600 dark:text-cyan-400 hover:underline disabled:opacity-50 disabled:cursor-not-allowed"
                                    disabled={isTuningScript}
                                >
                                    Upload from .txt / .docx...
                                </button>
                                <input 
                                    type="file" 
                                    ref={scriptTunerFileInputRef} 
                                    onChange={(e) => handleFileChange(e, setScriptTunerInput)} 
                                    accept=".txt,.docx" 
                                    className="hidden"
                                />
                            </div>
                        </div>
                        <div>
                            <label className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Select a Tone</label>
                            <select
                                value={selectedTone}
                                onChange={e => setSelectedTone(e.target.value)}
                                disabled={isTuningScript}
                                className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none transition-all duration-300"
                            >
                                {TONE_OPTIONS.map(tone => <option key={tone} value={tone}>{tone}</option>)}
                            </select>
                        </div>

                        {selectedTone === 'Custom' && (
                            <div className="animate-fade-in">
                                <label className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Describe the Custom Tone</label>
                                <input
                                    type="text"
                                    value={customTone}
                                    onChange={e => setCustomTone(e.target.value)}
                                    placeholder="e.g., A wise old wizard telling a story"
                                    className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none transition-all duration-300"
                                    disabled={isTuningScript}
                                />
                            </div>
                        )}
                        
                        <button
                            onClick={handleTuneScript}
                            disabled={isTuningScript || !scriptTunerInput.trim() || (selectedTone === 'Custom' && !customTone.trim())}
                            className="w-full mt-2 flex items-center justify-center gap-3 gradient-button text-white font-bold py-3 px-4 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-all"
                        >
                            {isTuningScript ? (
                                <>
                                    <div role="status" className="inline-block h-5 w-5 animate-spin rounded-full border-2 border-solid border-current border-r-transparent" />
                                    <span>Tuning Script...</span>
                                </>
                            ) : (
                                <>
                                    <TuneIcon className="h-5 w-5"/> Tune Script
                                </>
                            )}
                        </button>
                    </div>

                    {/* Output Column */}
                    <div className="space-y-4">
                        <div className="flex justify-between items-center">
                            <h3 className="text-lg font-semibold text-gray-800 dark:text-white">AI-Tuned Script (SSML)</h3>
                             <button
                                onClick={() => handleCopyToClipboard(scriptTunerOutput)}
                                disabled={!scriptTunerOutput || isTuningScript}
                                className="flex items-center gap-1.5 text-sm font-semibold bg-slate-200 dark:bg-[#0b1120] hover:bg-slate-300/80 dark:hover:bg-cyan-500/10 py-1.5 px-3 rounded-lg transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                            >
                                {isCopied ? <CheckCircleIcon className="h-5 w-5 text-green-500" /> : <ClipboardIcon className="h-5 w-5" />}
                                {isCopied ? 'Copied!' : 'Copy'}
                            </button>
                        </div>
                        <div className="w-full h-[360px] bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-3 text-base overflow-y-auto relative">
                           {isTuningScript ? (
                               <div className="absolute inset-0 flex items-center justify-center">
                                   <Spinner />
                               </div>
                           ) : scriptTunerOutput ? (
                                <pre className="whitespace-pre-wrap font-sans text-sm">{scriptTunerOutput}</pre>
                           ) : (
                                <div className="flex items-center justify-center h-full text-gray-500 dark:text-gray-400">
                                    Your enhanced script will appear here.
                                </div>
                           )}
                        </div>
                    </div>
                </div>
            </div>
        );
    }
    
    const renderBulkScriptWriter = () => {
        const canAnalyze = bulkScriptTitles.trim() && !isWorkingOnBulkScripts;
        const canGenerateManual = bulkScriptTitles.trim() && bulkScriptGuides.trim() && bulkScriptOutline.trim() && !isWorkingOnBulkScripts;
        
        return (
            <div className="space-y-6">
                 <p className="text-sm text-gray-600 dark:text-gray-400">
                    Generate multiple scripts at once. Use the manual mode if you have your own guides and outline, or use the AI-assisted mode to generate them from just titles.
                </p>

                <div className="flex space-x-1 bg-slate-100 dark:bg-[#0b1120] p-1 rounded-xl">
                    <button onClick={() => { setBulkScriptMode('manual'); setBulkScriptStep('input'); }} className={`w-full flex items-center justify-center gap-2 text-sm font-semibold py-1.5 px-2 rounded-lg transition-colors ${bulkScriptMode === 'manual' ? 'gradient-button text-white' : 'text-gray-600 dark:text-gray-300 hover:bg-slate-200 dark:hover:bg-cyan-500/10'}`}>
                        Manual Input
                    </button>
                    <button onClick={() => { setBulkScriptMode('auto-outline'); setBulkScriptStep('input'); }} className={`w-full flex items-center justify-center gap-2 text-sm font-semibold py-1.5 px-2 rounded-lg transition-colors ${bulkScriptMode === 'auto-outline' ? 'gradient-button text-white' : 'text-gray-600 dark:text-gray-300 hover:bg-slate-200 dark:hover:bg-cyan-500/10'}`}>
                        AI-Assisted Outline
                    </button>
                </div>
                
                <div className="p-4 bg-slate-100 dark:bg-[#0b1120] rounded-xl border border-slate-200 dark:border-[#252f41]">
                    <h4 className="text-base font-semibold text-gray-800 dark:text-white mb-3">API Configuration</h4>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div>
                            <label htmlFor="bulk-script-provider" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Provider</label>
                            <select
                                id="bulk-script-provider"
                                value={bulkScriptApiProvider}
                                onChange={e => handleBulkScriptApiProviderChange(e.target.value as BulkScriptApiProvider)}
                                className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none"
                            >
                                {BULK_SCRIPT_PROVIDERS.map(p => <option key={p.id} value={p.id}>{p.name}</option>)}
                            </select>
                        </div>
                        <div>
                            <label htmlFor="bulk-script-key" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">API Key</label>
                            <div className="flex items-center gap-2">
                                <input
                                    id="bulk-script-key"
                                    type="password"
                                    value={tempBulkScriptApiKey}
                                    onChange={e => setTempBulkScriptApiKey(e.target.value)}
                                    placeholder={bulkScriptApiProvider === 'gemini' ? 'Optional (uses app credits)' : 'Enter API Key'}
                                    className="flex-grow bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none disabled:opacity-50"
                                />
                                <button
                                    onClick={handleSaveBulkScriptKey}
                                    className="px-4 py-2.5 rounded-lg font-semibold transition-colors bg-slate-200 dark:bg-[#252f41] text-gray-800 dark:text-white hover:bg-slate-300 dark:hover:bg-cyan-500/10 disabled:opacity-50 disabled:cursor-not-allowed"
                                >
                                    Save
                                </button>
                            </div>
                        </div>
                        {bulkScriptApiProvider !== 'gemini' && tempBulkScriptApiKey && (
                            <div className="md:col-span-2 animate-fade-in">
                                <label htmlFor="bulk-script-model" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Model</label>
                                {isFetchingBulkScriptModels ? (
                                    <div className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-base text-gray-500">Fetching models...</div>
                                ) : bulkScriptModelError ? (
                                    <div className="w-full bg-red-100 dark:bg-red-900/20 border border-red-200 dark:border-red-900/30 rounded-lg p-2.5 text-sm text-red-700 dark:text-red-300">{bulkScriptModelError}</div>
                                ) : bulkScriptModels.length > 0 ? (
                                    <select
                                        id="bulk-script-model"
                                        value={bulkScriptSelectedModel}
                                        onChange={handleBulkScriptModelChange}
                                        className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none"
                                    >
                                        {bulkScriptModels.map(model => <option key={model} value={model}>{model}</option>)}
                                    </select>
                                ) : (
                                    <div className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2.5 text-base text-gray-500">No models found or key is invalid.</div>
                                )}
                            </div>
                        )}
                    </div>
                    {bulkScriptApiProvider !== 'gemini' && !tempBulkScriptApiKey &&
                        <p className="text-xs text-yellow-600 dark:text-yellow-400 mt-2">An API key is required to use the '{BULK_SCRIPT_PROVIDERS.find(p => p.id === bulkScriptApiProvider)?.name}' provider.</p>
                    }
                </div>

                {bulkScriptMode === 'manual' && (
                <div className="space-y-4 pt-4 border-t border-slate-200 dark:border-[#252f41] animate-fade-in">
                    <div>
                        <label htmlFor="bulk-titles" className="block text-base font-semibold text-gray-800 dark:text-white mb-2">1. Script Titles (one per line)</label>
                        <textarea id="bulk-titles" value={bulkScriptTitles} onChange={(e) => setBulkScriptTitles(e.target.value)} rows={5} className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-3 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none resize-y transition-colors" placeholder="e.g., The History of Ancient Rome&#10;The Future of Mars Colonization" disabled={isWorkingOnBulkScripts}/>
                        <div className="flex justify-end mt-1">
                            <button onClick={() => bulkScriptTitlesFileInputRef.current?.click()} className="text-sm font-medium text-cyan-600 dark:text-cyan-400 hover:underline">Upload Titles...</button>
                            <input type="file" ref={bulkScriptTitlesFileInputRef} onChange={(e) => handleFileChange(e, setBulkScriptTitles)} accept=".txt,.docx" className="hidden"/>
                        </div>
                    </div>
                    <div>
                        <label htmlFor="bulk-guides" className="block text-base font-semibold text-gray-800 dark:text-white mb-2">2. AI Writing Guides</label>
                        <textarea id="bulk-guides" value={bulkScriptGuides} onChange={(e) => setBulkScriptGuides(e.target.value)} rows={4} className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-3 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none resize-y transition-colors" placeholder="e.g., Write in an engaging, narrative style. Use simple language suitable for a wide audience. Start with a hook. The script should be around 700 words." disabled={isWorkingOnBulkScripts}/>
                         <div className="flex justify-end mt-1">
                            <button onClick={() => bulkScriptGuidesFileInputRef.current?.click()} className="text-sm font-medium text-cyan-600 dark:text-cyan-400 hover:underline">Upload Guides...</button>
                            <input type="file" ref={bulkScriptGuidesFileInputRef} onChange={(e) => handleFileChange(e, setBulkScriptGuides)} accept=".txt,.docx" className="hidden"/>
                        </div>
                    </div>
                    <div>
                        <label htmlFor="bulk-outline" className="block text-base font-semibold text-gray-800 dark:text-white mb-2">3. Script Template / Outline</label>
                        <textarea id="bulk-outline" value={bulkScriptOutline} onChange={(e) => setBulkScriptOutline(e.target.value)} rows={5} className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-3 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none resize-y transition-colors" placeholder={"Introduction:\n- Hook\n- Brief overview\n\nMain Body:\n- Point 1\n- Point 2\n- Point 3\n\nConclusion:\n- Summary\n- Call to action"} disabled={isWorkingOnBulkScripts}/>
                         <div className="flex justify-end mt-1">
                            <button onClick={() => bulkScriptOutlineFileInputRef.current?.click()} className="text-sm font-medium text-cyan-600 dark:text-cyan-400 hover:underline">Upload Outline...</button>
                            <input type="file" ref={bulkScriptOutlineFileInputRef} onChange={(e) => handleFileChange(e, setBulkScriptOutline)} accept=".txt,.docx" className="hidden"/>
                        </div>
                    </div>
                    <button onClick={handleGenerateBulkScripts} disabled={!canGenerateManual} className="w-full mt-4 flex items-center justify-center gap-3 gradient-button text-white font-bold py-3 px-4 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-all">
                        {isWorkingOnBulkScripts ? <BulkScriptLoadingIndicator /> : <><BulkScriptIcon/> <span>Generate Scripts</span></>}
                    </button>
                </div>
                )}

                {bulkScriptMode === 'auto-outline' && (
                    <div className="pt-4 border-t border-slate-200 dark:border-[#252f41] animate-fade-in">
                        <div className="space-y-4">
                            <div>
                                <label htmlFor="bulk-titles-auto" className="block text-base font-semibold text-gray-800 dark:text-white mb-2">1. Script Titles (one per line)</label>
                                <textarea id="bulk-titles-auto" value={bulkScriptTitles} onChange={(e) => setBulkScriptTitles(e.target.value)} rows={8} className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-3 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none resize-y" placeholder="e.g., The History of Ancient Rome&#10;The Future of Mars Colonization"/>
                                <div className="flex justify-end mt-1">
                                    <button onClick={() => bulkScriptTitlesFileInputRef.current?.click()} className="text-sm font-medium text-cyan-600 dark:text-cyan-400 hover:underline">Upload Titles...</button>
                                    <input type="file" ref={bulkScriptTitlesFileInputRef} onChange={(e) => handleFileChange(e, setBulkScriptTitles)} accept=".txt,.docx" className="hidden"/>
                                </div>
                            </div>
                            <div>
                                <label htmlFor="bulk-word-count-auto" className="block text-base font-semibold text-gray-800 dark:text-white mb-2">2. Default Target Word Count</label>
                                <input type="number" id="bulk-word-count-auto" value={bulkScriptWordCount} onChange={(e) => setBulkScriptWordCount(Number(e.target.value))} step={50} min={100} className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-3 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none"/>
                            </div>
                            <button onClick={handleAnalyzeTitles} disabled={!canAnalyze} className="w-full mt-4 flex items-center justify-center gap-3 gradient-button text-white font-bold py-3 px-4 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-all">
                                <WandIcon className="h-5 w-5"/> Analyze Titles & Create Outlines
                            </button>
                        </div>
                    </div>
                )}
            </div>
        );
    }

    const renderBulkScriptReviewer = () => {
        const canGenerateAuto = bulkScriptJobs.length > 0 && bulkScriptJobs.some(j => j.status === 'review') && !isWorkingOnBulkScripts;

        return (
            <div className="w-full animate-fade-in space-y-6">
                 <div className="flex justify-between items-center">
                    <h2 className="text-2xl font-bold text-gray-900 dark:text-white">Review & Approve Outlines</h2>
                    <button onClick={() => { setBulkScriptStep('input'); setBulkScriptJobs([]); }} className="text-sm font-medium text-cyan-600 dark:text-cyan-400 hover:underline">
                        Start Over
                    </button>
                </div>

                <div className="space-y-4 max-h-[calc(100vh-18rem)] overflow-y-auto pr-2">
                     {bulkScriptJobs.map(job => (
                        <div key={job.id} className="p-4 bg-white dark:bg-[#0b1120] rounded-xl border border-slate-200 dark:border-[#252f41] shadow-sm transition-all">
                             <div className="flex justify-between items-center">
                                <p className="font-bold text-lg text-gray-800 dark:text-white truncate" title={job.title}>{job.title}</p>
                                {job.status === 'analyzing' && <div role="status" className="w-5 h-5 animate-spin rounded-full border-2 border-solid border-current border-r-transparent text-cyan-500" />}
                                {job.status === 'error' && <span className="text-red-500 font-semibold">Error</span>}
                                {(job.status === 'review' || job.status === 'complete') && <CheckCircleIcon className="h-6 w-6 text-green-500" />}
                            </div>
                            <div className="mt-4 space-y-4 border-t border-slate-200 dark:border-[#252f41] pt-4 grid grid-cols-1 md:grid-cols-2 gap-6">
                                {job.status === 'error' && <p className="text-red-500 text-sm md:col-span-2">Error: {job.error}</p>}
                                {job.status === 'analyzing' && <div className="md:col-span-2 text-center py-10"><Spinner/></div>}
                                {(job.status === 'review' || job.status === 'complete') && (
                                    <>
                                        <div>
                                            <label className="block text-base font-semibold text-gray-700 dark:text-gray-300 mb-2">AI Writing Guides</label>
                                            <textarea value={job.guides} onChange={e => handleUpdateBulkJob(job.id, { guides: e.target.value })} rows={10} className="w-full text-sm bg-slate-100 dark:bg-[#040810] border border-slate-200 dark:border-[#252f41] rounded-lg p-2 focus:ring-1 focus:ring-cyan-500 outline-none resize-y" disabled={isWorkingOnBulkScripts}/>
                                        </div>
                                        <div>
                                            <label className="block text-base font-semibold text-gray-700 dark:text-gray-300 mb-2">Script Outline</label>
                                             <div className="w-full text-sm bg-slate-100 dark:bg-[#040810] border border-slate-200 dark:border-[#252f41] rounded-lg p-2 h-[230px] overflow-y-auto">
                                                <SimpleMarkdown text={job.outline || ''} />
                                             </div>
                                             {/* Keeping edit functionality for future if needed, but display-first is better */}
                                             {/* <textarea value={job.outline} onChange={e => handleUpdateBulkJob(job.id, { outline: e.target.value })} rows={10} className="w-full text-sm bg-slate-100 dark:bg-[#040810] border border-slate-200 dark:border-[#252f41] rounded-lg p-2 focus:ring-1 focus:ring-cyan-500 outline-none resize-y" disabled={isWorkingOnBulkScripts}/> */}
                                        </div>
                                    </>
                                )}
                            </div>
                        </div>
                    ))}
                </div>
                
                <button onClick={handleGenerateBulkScripts} disabled={!canGenerateAuto} className="w-full mt-4 flex items-center justify-center gap-3 gradient-button text-white font-bold py-3 px-4 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-all">
                    {isWorkingOnBulkScripts ? <BulkScriptLoadingIndicator /> : <><BulkScriptIcon/> <span>Approve & Generate All Scripts</span></>}
                </button>
            </div>
        )
    };

    const renderBulkScriptResults = () => {
        const generatedCount = bulkScriptJobs.filter(j => j.status === 'complete').length;
        const totalCount = bulkScriptJobs.length;
        const isWorking = isWorkingOnBulkScripts;

        return (
            <div className="space-y-4 h-full flex flex-col">
                <div className="flex-none">
                    <div className="flex flex-wrap gap-y-2 justify-between items-center">
                        <h2 className="text-xl font-bold text-gray-900 dark:text-white">Generated Scripts</h2>
                        {isWorking && (
                            <div className="flex items-center gap-2 text-sm text-cyan-600 dark:text-cyan-400 animate-fade-in">
                                <div role="status" className="w-4 h-4 animate-spin rounded-full border-2 border-solid border-current border-r-transparent" />
                                <span>{generatedCount} of {totalCount} Complete...</span>
                            </div>
                        )}
                        {!isWorking && totalCount > 0 &&(
                            <span className="text-sm font-medium text-gray-500 dark:text-gray-400">
                                {generatedCount} / {totalCount} generated
                            </span>
                        )}
                    </div>
                     {generatedCount > 0 && (
                        <div className="flex flex-wrap items-center gap-2 mt-4">
                             <button onClick={() => handleDownloadAllScripts('txt')} className="flex items-center gap-2 text-sm font-semibold bg-slate-200 dark:bg-[#172134] hover:bg-slate-300/80 dark:hover:bg-cyan-500/10 text-gray-700 dark:text-gray-300 py-1.5 px-3 rounded-lg transition-colors">
                                <ZipIcon /> Download All (.txt)
                            </button>
                             <button onClick={() => handleDownloadAllScripts('docx')} className="flex items-center gap-2 text-sm font-semibold bg-slate-200 dark:bg-[#172134] hover:bg-slate-300/80 dark:hover:bg-cyan-500/10 text-gray-700 dark:text-gray-300 py-1.5 px-3 rounded-lg transition-colors">
                                <ZipIcon /> Download All (.docx)
                            </button>
                        </div>
                    )}
                </div>
                
                <div className="space-y-4 flex-1 overflow-y-auto pr-2 -mr-2">
                    {totalCount === 0 && !isWorking && (
                        <div className="flex items-center justify-center h-full text-center text-gray-500 dark:text-gray-400 bg-slate-100 dark:bg-[#0b1120] rounded-xl border-2 border-dashed border-slate-300 dark:border-[#252f41] p-8">
                            <p>Your generated scripts will appear here.</p>
                        </div>
                    )}
                    {bulkScriptJobs.map(job => (
                        <div key={job.id} className="p-4 bg-white dark:bg-[#0b1120] rounded-xl border border-slate-200 dark:border-[#252f41] shadow-md transition-all">
                            <div className="flex justify-between items-start mb-2 gap-4">
                                <div className="flex-1">
                                    <h3 className="font-bold text-lg text-gray-800 dark:text-white truncate">{job.title}</h3>
                                    <StatusBadge status={job.status} />
                                </div>
                                <div className="flex items-center gap-2 flex-shrink-0">
                                    <button
                                        onClick={() => handleDownloadScriptAsTxt(job.script || '', job.title)}
                                        disabled={job.status !== 'complete'}
                                        className="p-2 bg-slate-100 dark:bg-slate-800/50 text-gray-600 dark:text-cyan-300 hover:bg-slate-200 dark:hover:bg-cyan-500/30 rounded-md transition-colors disabled:opacity-50"
                                        title="Download as .txt"
                                    >
                                        <TxtIcon className="h-4 w-4" />
                                    </button>
                                    <button
                                        onClick={() => handleDownloadScriptAsDocx(job.script || '', job.title)}
                                        disabled={job.status !== 'complete'}
                                        className="p-2 bg-slate-100 dark:bg-slate-800/50 text-gray-600 dark:text-cyan-300 hover:bg-slate-200 dark:hover:bg-cyan-500/30 rounded-md transition-colors disabled:opacity-50"
                                        title="Download as .docx"
                                    >
                                        <DocxIcon className="h-4 w-4" />
                                    </button>
                                </div>
                            </div>
                            
                            {(job.status === 'generating' || job.status === 'analyzing') && (
                                <div className="h-20 flex items-center justify-center text-gray-500 dark:text-gray-400">
                                    <Spinner />
                                </div>
                            )}
                            {job.status === 'error' && (
                                <div className="h-20 flex items-center justify-center text-red-500 bg-red-500/10 rounded-lg p-2 text-sm">
                                    <p>Error: {job.error}</p>
                                </div>
                            )}
                            {job.script && (
                                <details className="mt-3 group">
                                    <summary className="cursor-pointer text-sm font-medium text-cyan-600 dark:text-cyan-400 hover:underline list-none flex items-center gap-1">
                                        View Script
                                        <svg className="w-4 h-4 transition-transform group-open:rotate-90" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M9 5l7 7-7 7" /></svg>
                                    </summary>
                                    <pre className="whitespace-pre-wrap font-sans text-sm text-gray-700 dark:text-gray-300 bg-slate-50 dark:bg-[#172134] p-3 rounded-lg max-h-64 overflow-y-auto mt-2">
                                        {job.script}
                                    </pre>
                                </details>
                            )}
                        </div>
                    ))}
                </div>
            </div>
        );
    }

    const renderResultsArea = () => {
        const currentIsGenerating = isGenerating || isWeaving || isGeneratingScenes;
        const isPortrait = ratio === '9:16' || ratio === '3:4';

        return (
            <>
                {error && (
                    <div className="bg-red-100 dark:bg-red-900/30 border border-red-400 text-red-700 dark:text-red-300 px-4 py-3 rounded-xl relative mb-6" role="alert">
                        <strong className="font-bold">Error: </strong>
                        <span className="block sm:inline">{error}</span>
                    </div>
                )}

                {videoUrl && (
                    <div className="mb-8 animate-fade-in">
                        <h2 className="text-2xl font-bold text-gray-900 dark:text-white mb-4">Generated Video</h2>
                        <video src={videoUrl} controls className="w-full rounded-2xl shadow-lg border border-slate-200 dark:border-[#252f41]" />
                        <a href={videoUrl} download={`ai_video_${Date.now()}.webm`} className="mt-4 inline-flex items-center gap-2 gradient-button text-white font-semibold py-2 px-4 rounded-lg transition-transform hover:scale-105">
                            <DownloadIcon /> Download Video
                        </a>
                    </div>
                )}

                {(activeResults.length > 0 || isWeaving || currentIsGenerating) && (
                    <div className="mb-6 flex justify-between items-center">
                         <div className="flex items-center gap-4">
                            <h2 className="text-2xl font-bold text-gray-900 dark:text-white">Results</h2>
                             {currentIsGenerating && (
                                <div className="flex items-center gap-2 text-sm text-cyan-600 dark:text-cyan-400 animate-fade-in">
                                    <div role="status" className="w-4 h-4 animate-spin rounded-full border-2 border-solid border-current border-r-transparent" />
                                    <span>{`Generating ${progress} of ${activeResults.length}...`}</span>
                                </div>
                            )}
                        </div>
                        <div className="flex items-center gap-2">
                            <button onClick={downloadAll} disabled={activeResults.every(r => !r.imageUrl)} className="flex items-center gap-2 text-sm font-semibold bg-slate-200 dark:bg-[#172134] hover:bg-slate-300/80 dark:hover:bg-cyan-500/10 text-gray-700 dark:text-gray-300 py-1.5 px-3 rounded-lg transition-colors disabled:opacity-50 disabled:cursor-not-allowed">
                                <ZipIcon /> Download All
                            </button>
                            <button onClick={() => setActiveResults([])} className="flex items-center gap-2 text-sm font-semibold bg-red-100 dark:bg-red-900/30 text-red-700 dark:text-red-300 hover:bg-red-200 dark:hover:bg-red-900/50 py-1.5 px-3 rounded-lg transition-colors">
                                <TrashIcon /> Clear
                            </button>
                        </div>
                    </div>
                )}

                <div className={`grid gap-4 ${ratio === '9:16' ? 'grid-cols-2 sm:grid-cols-3 md:grid-cols-4' : 'grid-cols-1 sm:grid-cols-2'}`}>
                    {activeResults.map(job => (
                        <div key={job.id} className={`group relative overflow-hidden rounded-xl shadow-lg border border-slate-200 dark:border-[#252f41] bg-slate-200 dark:bg-[#172134] ${ratioMap[ratio]}`}>
                            {job.imageUrl ? (
                                <img src={job.imageUrl} alt={job.prompt} className="w-full h-full object-cover transition-transform group-hover:scale-105" />
                            ) : job.error ? (
                                <div className="flex flex-col items-center justify-center h-full p-4 text-center text-red-500 bg-red-100 dark:bg-red-900/20">
                                    <p className="font-semibold">Error</p>
                                    <p className="text-xs">{job.error}</p>
                                </div>
                            ) : (
                                <div className="w-full h-full flex items-center justify-center">
                                    <div role="status" className="w-8 h-8 animate-spin rounded-full border-4 border-solid border-cyan-400 border-r-transparent" />
                                </div>
                            )}
                             <div className={`absolute inset-0 bg-gradient-to-t from-black/80 via-black/40 to-transparent opacity-0 group-hover:opacity-100 transition-opacity duration-300 flex flex-col items-center justify-end text-center ${isPortrait ? 'p-2' : 'p-4'}`}>
                                <p className={`text-white font-semibold leading-tight drop-shadow-md line-clamp-3 ${isPortrait ? 'mb-3 text-sm' : 'mb-4 text-base'}`}>
                                    {job.prompt}
                                </p>
                                <div className={`flex items-center justify-center ${isPortrait ? 'gap-2' : 'gap-3'}`}>
                                    <button onClick={() => handleDownloadOne(job)} disabled={!job.imageUrl} className={`bg-white/20 backdrop-blur-sm rounded-full text-white hover:bg-white/30 transition-all disabled:opacity-50 disabled:cursor-not-allowed ${isPortrait ? 'p-2' : 'p-3'}`}>
                                        <DownloadIcon className={isPortrait ? 'h-4 w-4' : 'h-5 w-5'} />
                                    </button>
                                    <button onClick={() => setEditingJob(job)} disabled={!job.imageUrl} className={`bg-white text-black rounded-full shadow-lg hover:scale-110 transition-all disabled:opacity-50 disabled:cursor-not-allowed ${isPortrait ? 'p-3' : 'p-4'}`}>
                                        <EditIcon className={isPortrait ? 'h-5 w-5' : 'h-6 w-6'} />
                                    </button>
                                    <button onClick={() => handleRegenerateOne(job.id)} disabled={currentIsGenerating} className={`bg-white/20 backdrop-blur-sm rounded-full text-white hover:bg-white/30 transition-all disabled:opacity-50 disabled:cursor-not-allowed ${isPortrait ? 'p-2' : 'p-3'}`}>
                                        <RegenerateIcon className={isPortrait ? 'h-4 w-4' : 'h-5 w-5'} />
                                    </button>
                                </div>
                            </div>
                        </div>
                    ))}
                </div>
                
                {isDoneGenerating && activeResults.length > 0 && !videoUrl && (
                    <div className="mt-8 p-6 bg-slate-50 dark:bg-[#172134] border border-slate-200 dark:border-[#252f41] rounded-2xl animate-fade-in">
                        <div className="flex flex-col md:flex-row items-center justify-between gap-4">
                            <div>
                                <h3 className="text-xl font-bold text-gray-900 dark:text-white">Ready for the Next Step?</h3>
                                <p className="text-gray-600 dark:text-gray-400 mt-1">Turn your generated scenes into an animated video with captions and voiceover.</p>
                            </div>
                            <button onClick={() => setMode('story-weaver')} className="gradient-button font-semibold text-white py-3 px-6 rounded-full transition-transform hover:scale-105 flex items-center gap-2">
                                Create Video <RightArrowIcon/>
                            </button>
                        </div>
                    </div>
                )}

                {isDoneGenerating && mode === 'story-weaver' && renderVideoSettings()}
            </>
        )
    }

    const allCurrentKeys: AllApiKeys = {
        image: customImageApiKeys,
        voice: {
            elevenlabs: elevenLabsApiKey,
            wellsaid: wellSaidApiKey,
        },
        script: customScriptApiKeys,
    };

    return (
        <div className="min-h-screen">
            <div className="container mx-auto p-4 md:p-6 lg:p-8">
                 <header className="flex justify-between items-center py-4">
                    <div className="flex items-center gap-3 cursor-pointer" onClick={clearAll}>
                        <LogoIcon className="h-8 w-8" />
                        <h1 className="text-xl font-bold animated-title-gradient">Ai Studio By Mr Hassnain</h1>
                    </div>
                    <div className="flex items-center gap-2 sm:gap-4">
                         {hasStarted && (
                            <CreditDisplay
                                credits={credits}
                                nextReset={creditResetTime}
                                isCustomKey={!!customImageApiKeys.gemini}
                                apiProvider={apiProvider}
                            />
                        )}
                         <button
                            onClick={() => setIsApiKeyModalOpen(true)}
                            className={`p-2 rounded-full transition-colors relative ${hasAnyCustomKey ? 'bg-cyan-100 dark:bg-cyan-900/40 text-cyan-600 dark:text-cyan-300' : 'bg-slate-200 dark:bg-[#172134] hover:bg-slate-300 dark:hover:bg-cyan-500/10'}`}
                            title="Set Custom API Keys"
                            aria-label="Set Custom API Keys"
                        >
                            <KeyIcon className="h-5 w-5" />
                            {hasAnyCustomKey && <span className="absolute top-0 right-0 block h-2 w-2 rounded-full bg-green-500 ring-2 ring-slate-100 dark:ring-gray-800" />}
                        </button>
                        <div className="flex items-center gap-1 p-1 bg-slate-200 dark:bg-[#172134] rounded-full border border-slate-300 dark:border-[#252f41]">
                            <button
                                onClick={() => setTheme('light')}
                                className={`p-1.5 rounded-full transition-colors ${theme === 'light' ? 'bg-white text-cyan-600 shadow' : 'text-gray-400 hover:text-gray-800'}`}
                                aria-label="Switch to light theme"
                            >
                                <SunIcon className="h-4 w-4" />
                            </button>
                            <button
                                onClick={() => setTheme('dark')}
                                className={`p-1.5 rounded-full transition-colors ${theme === 'dark' ? 'bg-slate-700 text-white' : 'text-gray-500 hover:text-gray-900'}`}
                                aria-label="Switch to dark theme"
                            >
                                <MoonIcon className="h-4 w-4" />
                            </button>
                        </div>
                        <button onClick={handleStartCreating} className="gradient-button font-semibold text-white py-2 px-4 rounded-full transition-transform hover:scale-105 hidden sm:block">
                            Get Started
                        </button>
                    </div>
                </header>

                <main className={`py-8 ${hasStarted ? 'flex gap-8 items-start' : ''}`}>
                    {hasStarted ? (
                    <>
                        {/* Special case for bulk script review step to take over the screen */}
                        {(mode === 'bulk-script-writer' && bulkScriptStep === 'review') ? (
                            <div className="w-full">
                                {renderBulkScriptReviewer()}
                            </div>
                        ) : (
                        <>
                            {/* --- LEFT SIDE: CONTROLS --- */}
                            <div className="flex-none">
                                <div className="sticky top-8 flex items-start">
                                    <div className="flex flex-col items-center">
                                        {/* VERTICAL NAV */}
                                        <div className="w-auto bg-slate-50 dark:bg-[#172134] backdrop-blur-md flex flex-col items-center space-y-2 p-2 rounded-2xl border border-slate-200 dark:border-[#252f41] self-start">
                                            {TABS.map(tab => (
                                                <button
                                                    key={tab.key}
                                                    onClick={() => handleNavClick(tab.key)}
                                                    className={`w-32 flex flex-col items-center justify-center p-3 rounded-xl transition-all duration-200 ${(mode === tab.key && isPanelOpen) ? 'gradient-button text-white' : 'text-gray-500 dark:text-gray-400 hover:bg-slate-200 dark:hover:bg-cyan-500/10'}`}
                                                    title={tab.label}
                                                >
                                                    {React.cloneElement(tab.icon, { className: 'h-6 w-6' })}
                                                    <span className="text-sm font-semibold text-center mt-2 leading-tight">
                                                        {tab.label}
                                                    </span>
                                                </button>
                                            ))}
                                        </div>
                                    </div>

                                    {/* SLIDEOUT PANEL */}
                                    <div className={`transition-all duration-300 ease-in-out self-start ${isPanelOpen ? 'w-[840px] max-w-[calc(100vw-30rem)] ml-4' : 'w-0'}`}>
                                        <div className={`bg-slate-50 dark:bg-[#172134] rounded-2xl border border-slate-200 dark:border-[#252f41] h-full max-h-[calc(100vh-8rem)] overflow-y-auto ${isPanelOpen ? 'p-6' : 'p-0'}`}>
                                            {isPanelOpen && (
                                                <>
                                                    <div className="flex justify-between items-center mb-6">
                                                        <h2 className="text-xl font-bold text-gray-900 dark:text-white">
                                                            {TABS.find(t => t.key === mode)?.label}
                                                        </h2>
                                                        <button onClick={() => setIsPanelOpen(false)} className="p-1.5 rounded-full hover:bg-slate-200 dark:hover:bg-cyan-500/10 transition-colors">
                                                            <CloseIcon className="h-5 w-5 text-gray-800 dark:text-gray-200" />
                                                        </button>
                                                    </div>
                                                    
                                                    {mode === 'story-weaver' && (
                                                        <div className="space-y-4">
                                                            <p className="text-sm text-gray-600 dark:text-gray-400">Create a visual story. Start with an idea, your own script, or even your voice.</p>
                                                            
                                                            <div className="flex space-x-1 bg-slate-100 dark:bg-[#0b1120] p-1 rounded-xl">
                                                                {([
                                                                {key: 'idea', label: 'From Idea', icon: <LightbulbIcon/>},
                                                                {key: 'script', label: 'From Script', icon: <DocumentTextIcon/>},
                                                                {key: 'voice', label: 'From Voice', icon: <MicrophoneIcon/>},
                                                                ] as const).map(item => (
                                                                    <button 
                                                                        key={item.key} 
                                                                        onClick={() => setStoryWeaverMode(item.key)} 
                                                                        className={`w-full flex items-center justify-center gap-2 text-sm font-semibold py-1.5 px-2 rounded-lg transition-colors ${storyWeaverMode === item.key ? 'gradient-button text-white' : 'text-gray-600 dark:text-gray-300 hover:bg-slate-200 dark:hover:bg-cyan-500/10'}`}>
                                                                        {item.icon}
                                                                        {item.label}
                                                                    </button>
                                                                ))}
                                                            </div>

                                                            {storyWeaverMode === 'idea' && (
                                                                <div className="space-y-4 animate-fade-in">
                                                                    <div>
                                                                        <label htmlFor="story-idea" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Your Story Idea</label>
                                                                        <textarea id="story-idea" value={storyIdea} onChange={(e) => setStoryIdea(e.target.value)} rows={4} className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-3 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none transition-all duration-300 resize-y" placeholder="e.g., A robot explorer discovering a crystal cave on Mars" disabled={isWeaving}/>
                                                                    </div>
                                                                    <div>
                                                                        <label htmlFor="videoDuration" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Target Video Duration (minutes)</label>
                                                                        <input type="number" id="videoDuration" value={videoDuration} onChange={e => setVideoDuration(Math.max(1, Number(e.target.value)))} className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-3" disabled={isWeaving}/>
                                                                    </div>
                                                                </div>
                                                            )}

                                                            {storyWeaverMode === 'script' && (
                                                                <div className="space-y-4 animate-fade-in">
                                                                    <div>
                                                                        <label htmlFor="script-input-sw" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Your Script</label>
                                                                        <textarea id="script-input-sw" value={scriptInput} onChange={(e) => setScriptInput(e.target.value)} rows={8} className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-3 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none transition-all duration-300 resize-y" placeholder="Enter your story or script here..." disabled={isWeaving}/>
                                                                        <button onClick={() => scriptFileInputRef.current?.click()} className="text-sm mt-2 font-medium text-cyan-600 dark:text-cyan-400 hover:text-cyan-500 dark:hover:text-cyan-300">
                                                                            Or upload from .txt / .docx
                                                                        </button>
                                                                        <input type="file" ref={scriptFileInputRef} onChange={(e) => handleFileChange(e, setScriptInput)} accept=".txt,.docx" className="hidden" />
                                                                    </div>
                                                                    <div>
                                                                        <label htmlFor="numScenesStory" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Number of Scenes to Generate</label>
                                                                        <input 
                                                                            type="number" 
                                                                            id="numScenesStory" 
                                                                            min="1" 
                                                                            value={numScenes} 
                                                                            onChange={e => setNumScenes(Math.max(1, Number(e.target.value)))} 
                                                                            className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-3 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none transition-all duration-300" 
                                                                            placeholder="e.g., 8"
                                                                            disabled={isWeaving}
                                                                        />
                                                                    </div>
                                                                </div>
                                                            )}
                                                            
                                                            {storyWeaverMode === 'voice' && (
                                                                <div className="space-y-4 animate-fade-in text-center">
                                                                    <p className="text-sm text-gray-600 dark:text-gray-400">Upload an audio file. We'll transcribe it into a script for you.</p>
                                                                    <div className="flex items-center justify-center">
                                                                        <button onClick={() => audioFileInputRef.current?.click()} className="w-24 h-24 rounded-full flex items-center justify-center transition-all gradient-button shadow-lg hover:scale-105">
                                                                            <UploadIcon className="h-12 w-12 text-white"/>
                                                                        </button>
                                                                        <input type="file" ref={audioFileInputRef} onChange={handleAudioFileChange} accept="audio/*" className="hidden" />
                                                                    </div>
                                                                    {audioBlob && !isTranscribing && <p className="text-green-600 dark:text-green-400">Audio ready for transcription!</p>}
                                                                    {isTranscribing && <p className="text-cyan-600 dark:text-cyan-400">Transcribing audio...</p>}
                                                                    <button onClick={handleTranscribe} disabled={!audioBlob || isTranscribing || isWeaving} className="w-full mt-2 flex items-center justify-center gap-3 bg-cyan-600 hover:bg-cyan-700 text-white font-bold py-3 px-4 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-all">
                                                                        {isTranscribing ? 'Transcribing...' : 'Transcribe & Edit Script'}
                                                                    </button>
                                                                </div>
                                                            )}

                                                            {storyWeaverMode !== 'voice' && (
                                                                <>
                                                                    <div className="space-y-4 pt-4 border-t border-slate-200 dark:border-[#252f41]">
                                                                        <h3 className="text-base font-semibold text-gray-700 dark:text-white">Image Settings</h3>
                                                                        {renderConfigOptions(true)}
                                                                    </div>
                                                                    <button onClick={handleWeaveStory} disabled={isWeaving || isDescribingStyle || (storyWeaverMode === 'idea' && !storyIdea.trim()) || (storyWeaverMode === 'script' && !scriptInput.trim())} className="w-full mt-2 flex items-center justify-center gap-3 gradient-button text-white font-bold py-3 px-4 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-all">
                                                                        {isDescribingStyle ? (
                                                                            <span>Analyzing Style...</span>
                                                                        ) : isWeaving ? (
                                                                            <>
                                                                                <div role="status" className="inline-block h-5 w-5 animate-spin rounded-full border-2 border-solid border-current border-r-transparent motion-reduce:animate-[spin_1.5s_linear_infinite]" />
                                                                                <span>{`Creating... (${weavingStep})`}</span>
                                                                            </>
                                                                        ) : (
                                                                            <>
                                                                                <SparkleIcon/> Create Storyboard
                                                                            </>
                                                                        )}
                                                                    </button>
                                                                </>
                                                            )}
                                                        </div>
                                                    )}

                                                    {mode === 'bulk' && (
                                                        <div className="space-y-4">
                                                            <div>
                                                                <label htmlFor="prompts-input" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Enter Prompts (one per line)</label>
                                                                <textarea
                                                                    id="prompts-input"
                                                                    value={promptsInput}
                                                                    onChange={(e) => setPromptsInput(e.target.value)}
                                                                    rows={8}
                                                                    className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-3 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none transition-all duration-300 resize-y"
                                                                    placeholder="A futuristic cityscape at sunset&#10;A lone astronaut on a red planet&#10;An enchanted forest with glowing mushrooms"
                                                                    disabled={isGenerating || isDescribingStyle}
                                                                />
                                                            </div>
                                                            <div className="flex justify-center">
                                                                <button onClick={() => promptsFileInputRef.current?.click()} className="text-sm font-medium text-cyan-600 dark:text-cyan-400 hover:text-cyan-500 dark:hover:text-cyan-300">
                                                                    Or upload from .txt / .docx
                                                                </button>
                                                                <input type="file" ref={promptsFileInputRef} onChange={(e) => handleFileChange(e, setPromptsInput)} accept=".txt,.docx" className="hidden" />
                                                            </div>
                                                            <div className="space-y-4 pt-4 border-t border-slate-200 dark:border-[#252f41]">
                                                                <h3 className="text-base font-semibold text-gray-700 dark:text-white">Image Settings</h3>
                                                                {renderConfigOptions()}
                                                            </div>
                                                            <button onClick={handleGenerateClick} disabled={isGenerating || !promptsInput.trim() || isDescribingStyle} className="w-full mt-2 flex items-center justify-center gap-3 gradient-button text-white font-bold py-3 px-4 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-all">
                                                                {isDescribingStyle ? 'Analyzing Style...' : (isGenerating ? 'Generating...' : 'Generate Images')}
                                                            </button>
                                                        </div>
                                                    )}
                                                    {mode === 'script-to-scene' && (
                                                        <div className="space-y-4">
                                                            <div>
                                                                <label htmlFor="script-input" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Paste your script or story</label>
                                                                <textarea
                                                                    id="script-input"
                                                                    value={scriptInput}
                                                                    onChange={(e) => {
                                                                        setScriptInput(e.target.value);
                                                                        setExtractedCharacters([]);
                                                                        setMainCharacter(null);
                                                                        setCharacterDetails({});
                                                                    }}
                                                                    rows={8}
                                                                    className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-3 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none transition-all duration-300 resize-y"
                                                                    placeholder="Enter a story or script here..."
                                                                    disabled={isGeneratingScenes || isAnalyzingScript || isDescribingStyle}
                                                                />
                                                            </div>
                                                             <div className="flex justify-between items-center gap-2">
                                                                <button onClick={() => scriptFileInputRef.current?.click()} className="text-sm font-medium text-cyan-600 dark:text-cyan-400 hover:text-cyan-500 dark:hover:text-cyan-300">
                                                                    Or upload script...
                                                                </button>
                                                                <input type="file" ref={scriptFileInputRef} onChange={(e) => handleFileChange(e, setScriptInput)} accept=".txt,.docx" className="hidden" />
                                                                <button onClick={handleAnalyzeScript} disabled={isAnalyzingScript || !scriptInput.trim()} className="flex items-center gap-2 text-sm font-semibold bg-slate-200 dark:bg-[#0b1120] hover:bg-slate-300/80 dark:hover:bg-cyan-500/10 py-1.5 px-3 rounded-lg transition-colors disabled:opacity-50 disabled:cursor-not-allowed">
                                                                    {isAnalyzingScript ? <div className="h-4 w-4 animate-spin rounded-full border-2 border-solid border-current border-r-transparent" /> : <WandIcon className="h-4 w-4 text-cyan-500" />}
                                                                    {isAnalyzingScript ? 'Analyzing...' : 'Analyze Script'}
                                                                </button>
                                                            </div>
                                                            
                                                            {extractedCharacters.length > 0 && (
                                                                <div className="space-y-4 pt-4 border-t border-slate-200 dark:border-[#252f41] animate-fade-in">
                                                                    <h3 className="text-base font-semibold text-gray-700 dark:text-white">Define Characters</h3>
                                                                    <div className="space-y-5 max-h-[30vh] overflow-y-auto pr-2">
                                                                        {extractedCharacters.map(char => (
                                                                            <div key={char} className="p-3 bg-slate-100 dark:bg-[#0b1120] rounded-lg border border-slate-200 dark:border-[#252f41]">
                                                                                <div className="flex justify-between items-center mb-2">
                                                                                    <p className="font-bold text-gray-800 dark:text-white">{char}</p>
                                                                                    {mainCharacter === char && (
                                                                                        <span className="text-xs font-semibold bg-cyan-200 text-cyan-800 dark:bg-cyan-900/50 dark:text-cyan-300 px-2 py-0.5 rounded-full">
                                                                                            Main Character
                                                                                        </span>
                                                                                    )}
                                                                                </div>
                                                                                <textarea
                                                                                    value={characterDetails[char]?.description || ''}
                                                                                    onChange={e => setCharacterDetails(prev => ({ ...prev, [char]: { ...(prev[char] || { refImage: null }), description: e.target.value } }))}
                                                                                    rows={3}
                                                                                    className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-2 text-sm focus:ring-1 focus:ring-cyan-500 outline-none transition-colors"
                                                                                    placeholder={`e.g., A tall man with a scar over his left eye...`}
                                                                                    disabled={isDescribingImage === char}
                                                                                />
                                                                                <div className="mt-2 flex items-center gap-3">
                                                                                    <button
                                                                                        onClick={() => characterRefImageRefs.current[char]?.click()}
                                                                                        disabled={isDescribingImage === char}
                                                                                        className="flex-1 text-xs font-semibold bg-slate-100 dark:bg-[#0b1120] hover:bg-slate-200 dark:hover:bg-cyan-500/10 text-gray-700 dark:text-gray-200 py-1.5 px-2 rounded-lg border border-slate-300 dark:border-[#252f41] transition-colors disabled:opacity-50"
                                                                                    >
                                                                                        {isDescribingImage === char ? 'Analyzing...' : (characterDetails[char]?.refImage ? 'Change Ref' : 'Upload Ref')}
                                                                                    </button>
                                                                                    <input
                                                                                        type="file"
                                                                                        accept="image/*"
                                                                                        ref={el => { if (el) characterRefImageRefs.current[char] = el; }}
                                                                                        onChange={e => handleCharacterRefImageUpload(e, char)}
                                                                                        className="hidden"
                                                                                    />
                                                                                    {characterDetails[char]?.refImage && (
                                                                                        <div className="relative group">
                                                                                            <img src={characterDetails[char]!.refImage!.url} alt={`${char} reference`} className="w-10 h-10 rounded-md object-cover"/>
                                                                                            <button
                                                                                                onClick={() => setCharacterDetails(prev => ({...prev, [char]: { ...prev[char], refImage: null }}))}
                                                                                                className="absolute -top-1.5 -right-1.5 bg-red-500 text-white rounded-full p-0.5 opacity-0 group-hover:opacity-100 transition-opacity"
                                                                                                aria-label={`Remove ${char} reference`}
                                                                                            >
                                                                                                <CloseIcon className="h-3 w-3" />
                                                                                            </button>
                                                                                        </div>
                                                                                    )}
                                                                                </div>
                                                                            </div>
                                                                        ))}
                                                                    </div>
                                                                    <p className="text-xs text-gray-500 dark:text-gray-400">Descriptions will be used to generate consistent characters.</p>
                                                                </div>
                                                            )}

                                                            <div className="space-y-4 pt-4 border-t border-slate-200 dark:border-[#252f41]">
                                                                <div>
                                                                    <label htmlFor="numScenes" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Number of Scenes to Generate</label>
                                                                    <input 
                                                                        type="number" 
                                                                        id="numScenes" 
                                                                        min="1" 
                                                                        value={numScenes} 
                                                                        onChange={e => setNumScenes(Math.max(1, Number(e.target.value)))} 
                                                                        className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-3 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none transition-all duration-300" 
                                                                        placeholder="e.g., 5"
                                                                        disabled={isGeneratingScenes || isDescribingStyle}
                                                                    />
                                                                </div>
                                                                <h3 className="text-base font-semibold text-gray-700 dark:text-white">Image Settings</h3>
                                                                {renderConfigOptions(false)}
                                                            </div>

                                                            <button 
                                                                onClick={handleScriptToSceneGenerate} 
                                                                disabled={isGeneratingScenes || !scriptInput.trim() || isAnalyzingScript || isDescribingStyle}
                                                                className="w-full mt-2 flex items-center justify-center gap-3 gradient-button text-white font-bold py-3 px-4 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-all"
                                                            >
                                                                {isDescribingStyle ? 'Analyzing Style...' : (isGeneratingScenes ? 'Generating Scenes...' : 'Generate Scenes')}
                                                            </button>
                                                        </div>
                                                    )}
                                                    {mode === 'character' && (
                                                        <div className="space-y-4">
                                                            <p className="text-sm text-gray-600 dark:text-gray-400">Create a consistent character sheet for your stories.</p>
                                                            <div>
                                                                <label htmlFor="character-description" className="block text-sm font-medium text-gray-600 dark:text-gray-400 mb-2">Character Description</label>
                                                                <textarea
                                                                    id="character-description"
                                                                    value={characterDescription}
                                                                    onChange={(e) => setCharacterDescription(e.target.value)}
                                                                    rows={5}
                                                                    className="w-full bg-white dark:bg-[#0b1120] border border-slate-200 dark:border-[#252f41] rounded-lg p-3 text-base focus:ring-2 focus:ring-cyan-500/50 focus:border-cyan-500 outline-none transition-all duration-300 resize-y"
                                                                    placeholder="e.g., A brave knight with silver armor and a red cape, determined expression, short brown hair, standing in a field."
                                                                    disabled={isGeneratingCharacter}
                                                                />
                                                            </div>
                                                            <button onClick={generateCharacterSheet} disabled={isGeneratingCharacter || !characterDescription.trim()} className="w-full mt-2 flex items-center justify-center gap-3 gradient-button text-white font-bold py-3 px-4 rounded-lg disabled:opacity-50 disabled:cursor-not-allowed transition-all">
                                                                {isGeneratingCharacter ? 'Generating Sheet...' : 'Generate Character Sheet'}
                                                            </button>
                                                        </div>
                                                    )}
                                                    {mode === 'bulk-script-writer' && renderBulkScriptWriter()}
                                                    {mode === 'voice-generator' && renderVoiceGenerator()}
                                                    {mode === 'ai-script-tuner' && renderScriptTuner()}
                                                </>
                                            )}
                                        </div>
                                    </div>
                                </div>
                            </div>

                            {/* --- RIGHT SIDE: RESULTS / MAIN CONTENT --- */}
                            <div className="flex-1 min-w-0">
                               {mode !== 'voice-generator' && mode !== 'ai-script-tuner' && mode !== 'bulk-script-writer' && renderResultsArea()}

                               {mode === 'character' && (
                                <div className="mt-8">
                                    <h2 className="text-2xl font-bold text-gray-900 dark:text-white mb-4">Character Sheet</h2>
                                    {isGeneratingCharacter && characterSheet.length === 0 && <Spinner />}
                                    <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                                        {characterSheet.map(job => (
                                            <div key={job.id} onClick={() => job.imageUrl && setActiveCharacter(job)} className={`group relative overflow-hidden rounded-xl shadow-lg border-2 ${activeCharacter?.id === job.id ? 'border-cyan-500' : 'border-slate-200 dark:border-[#252f41]'} aspect-w-1 aspect-h-1 bg-slate-200 dark:bg-[#172134] cursor-pointer`}>
                                                {job.imageUrl ? (
                                                    <img src={job.imageUrl} alt={job.prompt} className="w-full h-full object-cover transition-transform group-hover:scale-105" />
                                                ) : (
                                                    <div className="flex items-center justify-center h-full">
                                                        <div role="status" className="w-8 h-8 animate-spin rounded-full border-4 border-solid border-cyan-400 border-r-transparent" />
                                                    </div>
                                                )}
                                                {activeCharacter?.id === job.id && (
                                                    <div className="absolute inset-0 bg-black/70 flex items-center justify-center">
                                                        <CheckCircleIcon className="h-12 w-12 text-green-400"/>
                                                    </div>
                                                )}
                                            </div>
                                        ))}
                                    </div>
                                </div>
                               )}

                               {mode === 'voice-generator' && voiceGeneratorMode === 'voice' && finalAudioFiles.length > 0 && renderVoiceGeneratorResultsPanel()}
                               {mode === 'bulk-script-writer' && (bulkScriptJobs.length > 0 || isWorkingOnBulkScripts) && (
                                    <div className="h-full">
                                        {renderBulkScriptResults()}
                                    </div>
                               )}
                            </div>
                        </>
                        )}
                    </>
                    ) : (
                        <div className="text-center py-20 animate-fade-in">
                            <div className="inline-block p-4 bg-white/50 dark:bg-black/20 backdrop-blur-md rounded-3xl border border-slate-200 dark:border-white/10 mb-8">
                                <LogoIcon className="h-16 w-16"/>
                            </div>
                            <h1 className="text-4xl md:text-6xl font-black text-gray-900 dark:text-white tracking-tighter">
                                Transform Your Ideas into <br/>
                                <span className="text-transparent bg-clip-text bg-gradient-to-r from-cyan-400 to-blue-600">Visual Reality</span>
                            </h1>
                            <p className="max-w-2xl mx-auto mt-6 text-lg text-gray-600 dark:text-gray-400">
                                Effortlessly generate stunning images, visualize stories, and create compelling videos from simple text. Your creative powerhouse is here.
                            </p>
                            <div className="mt-10">
                                <button onClick={handleStartCreating} className="gradient-button text-white font-bold py-4 px-8 rounded-full text-lg transition-transform hover:scale-105 shadow-lg">
                                    Start Creating Now
                                </button>
                            </div>
                        </div>
                    )}
                </main>
                
                {editingJob && (
                    <EditModal 
                        job={editingJob} 
                        apiProvider={apiProvider}
                        imageStyles={imageStyles}
                        aspectRatio={ratio}
                        steps={steps}
                        apiKey={customImageApiKeys[apiProvider] || null}
                        onClose={() => setEditingJob(null)}
                        onSave={handleSaveEdit}
                        onSpendCredit={spendCredits}
                    />
                )}
                 {isApiKeyModalOpen && (
                    <ApiKeyModal
                        isOpen={isApiKeyModalOpen}
                        onClose={() => setIsApiKeyModalOpen(false)}
                        currentKeys={allCurrentKeys}
                        onSave={handleSaveAllApiKeys}
                        onClear={handleClearAllApiKeys}
                        onSaveSingleKey={handleSaveSingleKey}
                    />
                )}
            </div>
        </div>
    );
};

export default App;
