import { TagsPrompt } from "@/constants/TagsPrompt";
import { OpenAICompletionOptions, OpenAIModelType } from "@/types";
import { UserMeetingRole } from "@microsoft/live-share";
import { useLiveCoPilot, useSharedState } from "@microsoft/live-share-react";
import { useEffect } from "react";
import { useGetCompletion } from "./useGetCompletion";

const ALLOWED_MEETING_ROLES = [
    UserMeetingRole.organizer,
    UserMeetingRole.presenter,
];
const DEFAULT_PROMPT_VALUE = "";
const AUTO_COMPLETIONS_ENABLED = true;
const DEFAULT_COMPLETIONS_DEBOUNCE_DELAY_MILLISECONDS = 2500;
const LOCK_PROMPT = true;
const LOCK_COMPLETION = true;
const OPEN_AI_MODEL_TYPE = OpenAIModelType.curie001;
const OPEN_AI_OPTIONS: OpenAICompletionOptions = {
    temperature: 0.0,
};

export const useIdeaTags = (ideaBoardId: string, promptText: string) => {
    const [ideaTags, setIdeaTags] = useSharedState<string[]>(`${ideaBoardId}-tags`, []);

    const onGetCompletion = useGetCompletion(
        OPEN_AI_MODEL_TYPE,
        OPEN_AI_OPTIONS
    );
    const {
        liveCoPilot,
        completionValue,
        changePrompt,
    } = useLiveCoPilot(
        `${ideaBoardId}-tags`,
        onGetCompletion,
        ALLOWED_MEETING_ROLES,
        DEFAULT_PROMPT_VALUE,
        AUTO_COMPLETIONS_ENABLED,
        DEFAULT_COMPLETIONS_DEBOUNCE_DELAY_MILLISECONDS,
        LOCK_PROMPT,
        LOCK_COMPLETION
    );

    // Update `useLiveCoPilot` prompt when `promptText` changes
    useEffect(() => {
        if (promptText.length < 1 || !liveCoPilot?.haveCompletionLock) return;
        const fullGeneratePrompt = `${TagsPrompt}\nHUMAN: ${promptText}\nTAGS:`;
        changePrompt(fullGeneratePrompt);
    }, [promptText, changePrompt, liveCoPilot]);

    useEffect(() => {
        if (typeof completionValue === "string") {
            const trimmedResponseText = completionValue.trimStart();
            const newTags = trimmedResponseText
                .split(", ")
                .map((t) => t.trim())
                .filter((t) => !!t);
            setIdeaTags(newTags);
        }
    }, [completionValue]);

    return {
        ideaTags,
    };
}