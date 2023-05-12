/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { LiveShareClient, TestLiveShareHost } from "@microsoft/live-share";
import { InsecureTokenProvider } from "@fluidframework/test-client-utils";
import { LiveCanvas } from "@microsoft/live-share-canvas";
import { LiveMediaSession } from "@microsoft/live-share-media";
import { SharedMap } from "fluid-framework";
import { useEffect, useRef, useState } from "react";
import { LiveEvent, LivePresence } from "@microsoft/live-share";
import { mediaList } from "../utils/media-list";
import { LiveShareHost } from "@microsoft/teams-js";

/**
 * Hook that creates/loads the apps shared objects.
 *
 * @remarks
 * This is an application specific hook that defines the fluid schema of Distributed Data Structures (DDS)
 * used by the app and passes that schema to the `LiveShareClient` to create/load your Fluid container.
 *
 * @returns Shared objects managed by the apps fluid container.
 */
export function useSharedObjects() {
    const startedRef = useRef(false);
    const [results, setResults] = useState();
    const [error, setError] = useState();

    useEffect(() => {
        // This hook should only be called once, so we use a ref to track if it has been called.
        // This is a workaround for the fact that useEffect is called twice on initial render in React V18.
        // In production, you might consider using React Suspense if you are using React V18.
        // We are not doing this here because many customers are still using React V17.
        // We are monitoring the React Suspense situation closely and may revisit in the future.
        if (startedRef.current) return;
        startedRef.current = true;
        console.log("useSharedObjects: starting");
        // Check if user is in Teams
        const url = window.location.href.includes("/#/")
            ? new URL(`${window.location.href.split("/#/").join("/")}`)
            : new URL(window.location);
        const inTeams = !!url.searchParams.get("inTeams");

        // Enable debugger
        window.localStorage.debug = "fluid:*";

        // Define container callback (optional).
        // * This is only called once when the container is first created.
        const onFirstInitialize = (container) => {
            console.log("useSharedObjects: onFirstInitialize called");
            // Setup any initial state here
            mediaList.forEach((mediaItem) => {
                container.initialObjects.playlistMap.set(mediaItem.id, {
                    ...mediaItem,
                    timeAdded: 0,
                });
            });
            container.initialObjects.playlistMap.set(
                "selected-media-id",
                mediaList[0].id
            );
        };

        // Define container schema
        const schema = {
            initialObjects: {
                presence: LivePresence,
                mediaSession: LiveMediaSession,
                notificationEvent: LiveEvent,
                liveCanvas: LiveCanvas,
                takeControlMap: SharedMap,
                playlistMap: SharedMap,
            },
        };

        // Create live share host
        const host = inTeams
            ? LiveShareHost.create()
            : TestLiveShareHost.create();

        // Create the client, join container, and set results
        console.log("useSharedObjects: joining container");
        const client = new LiveShareClient(host);
        client
            .joinContainer(schema, onFirstInitialize)
            .then((results) => {
                console.log("useSharedObjects: joined container");
                setResults(results);
            })
            .catch((err) => setError(err));
    }, []);

    const container = results?.container;
    const initialObjects = container?.initialObjects;
    return {
        presence: initialObjects?.presence,
        mediaSession: initialObjects?.mediaSession,
        notificationEvent: initialObjects?.notificationEvent,
        takeControlMap: initialObjects?.takeControlMap,
        playlistMap: initialObjects?.playlistMap,
        inkEvent: initialObjects?.inkEvent,
        liveCanvas: initialObjects?.liveCanvas,
        container,
        error,
        services: results?.services,
        timestampProvider: results?.timestampProvider,
    };
}
