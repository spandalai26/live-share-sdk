/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the Microsoft Live Share SDK License.
 */

import {
    IFluidContainer,
    LoadableObjectClass,
    LoadableObjectClassRecord,
    SharedMap,
} from "fluid-framework";
import { AzureContainerServices } from "@fluidframework/azure-client";
import { IFluidLoadable, FluidObject } from "@fluidframework/core-interfaces";
import { IFluidTurboClient } from "./interfaces/IFluidTurboClient";
import { DynamicObjectManager } from "./dds-objects";
import { DynamicObjectRegistry } from "@microsoft/live-share";
import { BASE_CONTAINER_SCHEMA } from "./internals";

/**
 * Base class for building Fluid Turbo clients.
 * @remarks
 * Unlike other Fluid clients, the turbo client wraps functionality regularly exposed through an `IFluidContainer`. This is due to the more opinionated
 * nature of this package than vanilla Fluid Framework, where developers do not define a full `ContainerSchema` themselves and objects are loaded dynamically.
 */
export abstract class FluidTurboClient implements IFluidTurboClient {
    private _awaitConnectedPromise?: Promise<void>;

    /**
     * Get the Fluid join container results
     */
    public abstract get results():
        | {
              container: IFluidContainer;
              services: AzureContainerServices;
          }
        | undefined;

    /**
     * @see IFluidTurboClient.stateMap
     */
    public get stateMap(): SharedMap | undefined {
        if (this.results) {
            return this.results.container.initialObjects
                .TURBO_STATE_MAP as SharedMap;
        }
        return undefined;
    }

    private get dynamicObjects(): DynamicObjectManager | undefined {
        if (this.results) {
            return this.results.container.initialObjects
                .TURBO_DYNAMIC_OBJECTS as DynamicObjectManager;
        }
        return undefined;
    }

    /**
     * @see IFluidTurboClient.getDDS
     */
    public async getDDS<
        T extends IFluidLoadable = FluidObject<any> & IFluidLoadable
    >(
        objectKey: string,
        objectClass: LoadableObjectClass<T>,
        onDidFirstInitialize?: (dds: T) => void
    ): Promise<T> {
        // The uniqueKey key makes the developer provided uniqueKey never conflict across different DDS objects
        if (!this.results) {
            throw new Error(
                "FluidTurboClient getDDS: cannot call until successful get/create/join FluidContainer"
            );
        }
        if (!this.dynamicObjects) {
            throw new Error(
                "FluidTurboClient: getDDS must have valid dynamicObjects DynamicObjectManager"
            );
        }
        await this.waitUntilConnected();
        const initialDDS = this.results.container.initialObjects[objectKey] as
            | T
            | undefined;
        if (initialDDS !== undefined) {
            return initialDDS;
        }
        const uniqueKey = `<${objectClass.name}>:${objectKey}`;
        const response = await this.dynamicObjects.getDDS<T>(
            uniqueKey,
            objectClass,
            this.results.container
        );
        if (response.created) {
            onDidFirstInitialize?.(response.dds);
        }
        return response.dds;
    }

    /**
     * Get the container schema to use within a `FluidTurboClient` container.
     *
     * @param initialObjects Optional. Initial objects to add to the schema
     * @returns a `ContainerSchema` record to use in a Fluid container
     */
    protected getContainerSchema(initialObjects?: LoadableObjectClassRecord) {
        return {
            initialObjects: {
                ...BASE_CONTAINER_SCHEMA.initialObjects,
                ...initialObjects,
            },
            // Get the static registry of LoadableObjectClass types.
            dynamicObjectTypes: [
                ...DynamicObjectRegistry.dynamicLoadableObjects.values(),
            ],
        };
    }

    private async waitUntilConnected(): Promise<void> {
        if (this._awaitConnectedPromise) {
            return this._awaitConnectedPromise;
        }
        this._awaitConnectedPromise = new Promise((resolve, reject) => {
            if (!this.results?.container) {
                reject(
                    new Error(
                        "FluidTurboClient awaitConnected: cannot load DDS without a Fluid container"
                    )
                );
                this._awaitConnectedPromise = undefined;
            } else {
                const onConnected = () => {
                    this.results?.container.off("connected", onConnected);
                    resolve();
                };
                // Wait until connected event to ensure we have the latest document
                // and don't accidentally override a dds handle recently created
                // by another client
                if (this.results.container.connectionState === 2) {
                    resolve();
                } else {
                    this.results.container.on("connected", onConnected);
                }
            }
        });
        return this._awaitConnectedPromise;
    }
}
