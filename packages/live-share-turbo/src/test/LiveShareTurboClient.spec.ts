/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the Microsoft Live Share SDK License.
 */

import { strict as assert } from "assert";
import {
    IFluidContainer,
    LoadableObjectClassRecord,
    SharedMap,
} from "fluid-framework";
import { LiveEvent, TestLiveShareHost } from "@microsoft/live-share";
import { LiveShareTurboClient } from "../LiveShareTurboClient";
import { AzureContainerServices } from "@fluidframework/azure-client";

describe("LiveShareTurboClient", () => {
    (window.performance as any).mark = () => {
        return {};
    };
    (window.performance as any).measure = () => {
        return {};
    };

    let containerId: string | undefined;
    const getContainerId = (): string | undefined => {
        return containerId;
    };
    const setContainerId = (newContainerId: string) => {
        containerId = newContainerId;
    };
    const host = TestLiveShareHost.create(getContainerId, setContainerId);
    const client1 = new LiveShareTurboClient(host);
    const client2 = new LiveShareTurboClient(host);

    const testMapKey = "TEST-MAP-KEY";
    const testLiveEventKey = "TEST-LIVE-EVENT-KEY";
    let results1: {
        container: IFluidContainer;
        services: AzureContainerServices;
        created: boolean;
    };
    let results2: {
        container: IFluidContainer;
        services: AzureContainerServices;
        created: boolean;
    };

    beforeEach(async () => {
        const initialObjects: LoadableObjectClassRecord = {
            [testLiveEventKey]: LiveEvent,
        };
        containerId = undefined;
        results1 = await client1.join(initialObjects);
        results2 = await client2.join(initialObjects);
    });

    it("Containers should be configured correctly", async () => {
        assert(
            [results1.created, results2.created].filter((created) => created)
                .length === 1,
            "Incorrect number of containers created"
        );
        assert(
            !!client1.results && !!client2.results,
            "client.results results not defined"
        );
        assert(
            !!results1.container || !!results1.services,
            "client1 results container or services are not defined"
        );
        assert(
            !!client1.stateMap || !!client2.stateMap,
            "stateMap is not defined"
        );
    });

    it("Should create dds from one client and join from another, correctly handle initial objects", async () => {
        let object1Created = false;
        const promise1 = client1.getDDS<SharedMap>(
            testMapKey,
            SharedMap,
            (dds: SharedMap) => {
                object1Created = true;
                assert(
                    dds !== undefined,
                    "dds is not defined in onFirstInitialize callback"
                );
            }
        );
        let object2Created = false;
        const promise2 = client2.getDDS<SharedMap>(
            testMapKey,
            SharedMap,
            (dds: SharedMap) => {
                object2Created = true;
                assert(
                    dds !== undefined,
                    "dds is not defined in onFirstInitialize callback"
                );
            }
        );

        // Wait for dds to to be created
        const [dds1, dds2] = await Promise.all([promise1, promise2]);

        assert(
            dds1 !== undefined && dds2 !== undefined,
            "test map(s) not defined"
        );
        // Only one should be marked as created
        assert(
            [object1Created, object2Created].filter((created) => created)
                .length === 1,
            "Incorrect number of objects created"
        );
    });

    it("Should not dynamically load initial objects", async () => {
        let liveEvent1Created = false;
        const liveEventPromise1 = client1.getDDS<LiveEvent>(
            testLiveEventKey,
            LiveEvent,
            () => {
                liveEvent1Created = true;
            }
        );
        let liveEvent2Created = false;
        const liveEventPromise2 = client2.getDDS<LiveEvent>(
            testLiveEventKey,
            LiveEvent,
            () => {
                liveEvent2Created = true;
            }
        );

        // Wait for get dds
        const [liveEvent1, liveEvent2] = await Promise.all([
            liveEventPromise1,
            liveEventPromise2,
        ]);
        assert(
            liveEvent1 !== undefined && liveEvent2 !== undefined,
            "test map(s) not defined"
        );
        // Should be zero because they are created as initialObjects, not through the DynamicObjectManager
        assert(
            [liveEvent1Created, liveEvent2Created].filter((created) => created)
                .length === 0,
            "Incorrect number of liveEvent objects marked as created"
        );
    });
});
