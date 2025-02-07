/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the Microsoft Live Share SDK License.
 */

import { strict as assert } from "assert";
import { LiveEventScope } from "../LiveEventScope";
import { LiveEventTarget } from "../LiveEventTarget";
import { LiveEventTimer } from "../LiveEventTimer";
import { MockRuntimeSignaler } from "./MockRuntimeSignaler";
import { LiveShareRuntime } from "../LiveShareRuntime";
import { TestLiveShareHost } from "../TestLiveShareHost";
import { LocalTimestampProvider } from "../LocalTimestampProvider";

function createConnectedSignalers() {
    const localRuntime = new MockRuntimeSignaler();
    const remoteRuntime = new MockRuntimeSignaler();
    MockRuntimeSignaler.connectRuntimes([localRuntime, remoteRuntime]);
    return { localRuntime, remoteRuntime };
}

describe("LiveEventTimer", () => {
    let localLiveRuntime = new LiveShareRuntime(
        TestLiveShareHost.create(),
        new LocalTimestampProvider()
    );
    let remoteLiveRuntime = new LiveShareRuntime(
        TestLiveShareHost.create(),
        new LocalTimestampProvider()
    );

    afterEach(async () => {
        // restore defaults
        localLiveRuntime = new LiveShareRuntime(
            TestLiveShareHost.create(),
            new LocalTimestampProvider()
        );
        remoteLiveRuntime = new LiveShareRuntime(
            TestLiveShareHost.create(),
            new LocalTimestampProvider()
        );
    });

    it("Should send a single event after a delay", (done) => {
        let created = 0;
        let triggered = 0;
        const signalers = createConnectedSignalers();
        const localScope = new LiveEventScope(
            signalers.localRuntime,
            localLiveRuntime
        );
        const localTarget = new LiveEventTarget(
            localScope,
            "test",
            (evt, local) => triggered++
        );
        const localTimer = new LiveEventTimer(
            localTarget,
            () => {
                created++;
                return {};
            },
            10
        );

        const remoteScope = new LiveEventScope(
            signalers.remoteRuntime,
            remoteLiveRuntime
        );
        const remoteTarget = new LiveEventTarget(
            remoteScope,
            "test",
            (evt, local) => triggered++
        );

        localTimer.start();
        assert(created == 0);
        setTimeout(() => {
            assert(created == 1, `Message creation count is ${created}`);
            assert(
                triggered == created * 2,
                `Messages created is ${created} but received is an unexpected ${triggered}`
            );
            localTimer.stop();
            done();
        }, 50);
    });

    it("Should repeatedly send events after a delay", (done) => {
        let created = 0;
        let triggered = 0;
        const signalers = createConnectedSignalers();
        const localScope = new LiveEventScope(
            signalers.localRuntime,
            localLiveRuntime
        );
        const localTarget = new LiveEventTarget(
            localScope,
            "test",
            (evt, local) => triggered++
        );
        const localTimer = new LiveEventTimer(
            localTarget,
            () => {
                created++;
                return {};
            },
            5,
            true
        );

        const remoteScope = new LiveEventScope(
            signalers.remoteRuntime,
            remoteLiveRuntime
        );
        const remoteTarget = new LiveEventTarget(
            remoteScope,
            "test",
            (evt, local) => triggered++
        );

        localTimer.start();
        assert(created == 0);
        setTimeout(() => {
            assert(created > 1, `Message creation count is ${created}`);
            assert(
                triggered == created * 2,
                `Messages created is ${created} but received is an unexpected ${triggered}`
            );
            localTimer.stop();
            done();
        }, 50);
    });
});
