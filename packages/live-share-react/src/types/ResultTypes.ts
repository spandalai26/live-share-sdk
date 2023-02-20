/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the Microsoft Live Share SDK License.
 */

import { AzureContainerServices } from "@fluidframework/azure-client";
import {
    ITimerConfig,
    LivePresenceUser,
    LiveEvent,
    LivePresence,
    LiveTimer,
} from "@microsoft/live-share";
import { InkingManager, LiveCanvas } from "@microsoft/live-share-canvas";
import {
    CoordinationWaitPoint,
    ExtendedMediaMetadata,
    MediaPlayerSynchronizer,
} from "@microsoft/live-share-media";
import { LiveAICompletion } from "@microsoft/live-share-turbo";
import { IFluidContainer, SharedMap } from "fluid-framework";
import { IReceiveLiveEvent } from "../interfaces";
import {
    OnAIChangeCompletionAction,
    OnAIChangePromptAction,
    OnPauseTimerAction,
    OnPlayTimerAction,
    OnStartTimerAction,
    OnUpdateLivePresenceAction,
    SendLiveEventAction,
} from "./ActionTypes";

export interface IAzureContainerResults {
    /**
     * Fluid Container.
     */
    container: IFluidContainer;
    /**
     * Azure container services which has information such as current socket connections.
     */
    services: AzureContainerServices;
}

export interface ILiveShareContainerResults extends IAzureContainerResults {
    /**
     * Whether the local user/client initially created the container.
     */
    created: boolean;
}

export interface IUseSharedMapResults<TData> {
    /**
     * Stateful map of most recent values from `SharedMap`.
     */
    map: ReadonlyMap<string, TData>;
    /**
     * Callback method to set/replace new entries in the `SharedMap`.
     */
    setEntry: (key: string, value: TData) => void;
    /**
     * Callback method to delete an existing entry in the `SharedMap`.
     */
    deleteEntry: (key: string) => void;
    /**
     * The Fluid `SharedMap` object, should you want to use it directly.
     */
    sharedMap: SharedMap | undefined;
}

export interface IUseLiveEventResults<TEvent extends object = object> {
    /**
     * The most recent event that has been received in the session.
     */
    latestEvent: IReceiveLiveEvent<TEvent> | undefined;
    /**
     * All received events since initializing this component, sorted from oldest -> newest.
     */
    allEvents: IReceiveLiveEvent<TEvent>[];
    /**
     * Callback method to send a new event to users in the session.
     */
    sendEvent: SendLiveEventAction<TEvent>;
    /**
     * The `LiveEvent` object, should you want to use it directly.
     */
    liveEvent: LiveEvent | undefined;
}

export interface IUseLiveTimerResults {
    /**
     * The current timer configuration.
     */
    timerConfig: ITimerConfig | undefined;
    /**
     * The time remaining in milliseconds.
     */
    milliRemaining: number | undefined;
    /**
     * The `LiveTimer` object, should you want to use it directly.
     */
    liveTimer: LiveTimer | undefined;
    /**
     * Start timer callback.
     */
    start: OnStartTimerAction;
    /**
     * Play timer callback.
     */
    play: OnPlayTimerAction;
    /**
     * Pause timer callback.
     */
    pause: OnPauseTimerAction;
}

export interface IUseLivePresenceResults<TData extends object = object> {
    /**
     * The local user's presence object.
     */
    localUser: LivePresenceUser<TData> | undefined;
    /**
     * List of non-local user's presence objects.
     */
    otherUsers: LivePresenceUser<TData>[];
    /**
     * List of all user's presence objects.
     */
    allUsers: LivePresenceUser<TData>[];
    /**
     * Live Share `LivePresence` object, should you want to use it directly.
     */
    livePresence: LivePresence<TData> | undefined;
    /**
     * Callback method to update the local user's presence.
     */
    updatePresence: OnUpdateLivePresenceAction<TData>;
}

export interface IUseMediaSynchronizerResults {
    /**
     * Stateful boolean on whether the session has an active suspension.
     */
    suspended: boolean;
    /**
     * Callback to initiate a play action for the group media session.
     */
    play: () => void;
    /**
     * Callback to initiate a pause action for the group media session.
     */
    pause: () => void;
    /**
     * Callback to initiate a seek action for the group media session.
     *
     * @param time timestamp of the video in seconds to seek to
     */
    seekTo: (time: number) => void;
    /**
     * Callback to change the track for the group media session.
     * @param track media metadata object, track src string, or null
     */
    setTrack: (track: Partial<ExtendedMediaMetadata> | string | null) => void;
    /**
     * Begin a new suspension. If a wait point is not set, the suspension will only impact the
     * local user.
     *
     * @param waitPoint Optional. Point in track to set the suspension at.
     * @see CoordinationWaitPoint
     */
    beginSuspension: (waitPoint?: CoordinationWaitPoint) => void;
    /**
     * End the currently active exception.
     */
    endSuspension: () => void;
    /**
     * Live Share `MediaPlayerSynchronizer` object, should you want to use it directly.
     */
    mediaSynchronizer: MediaPlayerSynchronizer | undefined;
}

export interface IUseLiveCanvasResults {
    /**
     * Inking manager for the canvas, which is synchronized by `liveCanvas`. It is set when loaded and undefined while loading.
     */
    inkingManager: InkingManager | undefined;
    /**
     * Live Canvas data object which synchronizes inking & cursors in the session. It is set when loaded and undefined while loading.
     */
    liveCanvas: LiveCanvas | undefined;
}

export interface IUseLiveAICompletionResults {
    /**
     * Stateful prompt value that will be used to generate completions.
     */
    promptValue: string;
    /**
     * Stateful completion value that was generated from the prompt value. Will be reset when `promptValue` is updated.
     */
    completionValue: string | undefined;
    /**
     * Local user is assigned to generate completions.
     */
    haveCompletionLock: boolean;
    /**
     * Callback method to update the prompt value.
     */
    changePrompt: OnAIChangePromptAction;
    /**
     * Callback method to send a completion value to users in the session.
     */
    sendCompletion: OnAIChangeCompletionAction;
    /**
     * The `LiveAICompletion` object, should you want to use it directly.
     */
    liveAICompletion: LiveAICompletion | undefined;
}
