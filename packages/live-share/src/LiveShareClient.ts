/*!
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the Microsoft Live Share SDK License.
 */

import {
    LiveShareHostDecorator,
    LiveShareTokenProvider,
    PolyfillHostDecorator,
    RoleVerifier,
} from "./internals";
import {
    AzureClient,
    AzureContainerServices,
    AzureLocalConnectionConfig,
    AzureRemoteConnectionConfig,
    ITelemetryBaseLogger,
    IUser,
} from "@fluidframework/azure-client";
import { ContainerSchema, IFluidContainer } from "@fluidframework/fluid-static";
import {
    ILiveShareHost,
    ContainerState,
    ITimestampProvider,
    IClientInfo,
    IRoleVerifier,
    UserMeetingRole,
} from "./interfaces";
import { HostTimestampProvider } from "./HostTimestampProvider";
import { InsecureTokenProvider } from "@fluidframework/test-client-utils";
import { TimestampProvider } from "./TimestampProvider";
import { LocalTimestampProvider } from "./LocalTimestampProvider";
import { TestLiveShareHost } from "./TestLiveShareHost";

/**
 * @hidden
 * Map v0.59 orderer endpoints to new v1.0 service endpoints
 */
const serviceEndpointMap = new Map<string | undefined, string>()
    .set(
        "https://alfred.westus2.fluidrelay.azure.com",
        "https://us.fluidrelay.azure.com"
    )
    .set(
        "https://alfred.westeurope.fluidrelay.azure.com",
        "https://eu.fluidrelay.azure.com"
    )
    .set(
        "https://alfred.southeastasia.fluidrelay.azure.com",
        "https://global.fluidrelay.azure.com"
    );

/**
 * Key for window global reference to loadable objects.
 * TODO: replace window reference with a better static registry system. Originally tried DynamicObjectRegistry._dynamicLoadableObjects
 * variable but that didn't work when testing locally, since all local package builds have separate symlink instances and thus have separate
 * static DynamicObjectRegistry classes.
 */
const GLOBAL_LIVE_SHARE_HOST_WINDOW_KEY = "@microsoft/live-share:LiveShareHost";
const GLOBAL_LIVE_SHARE_TIMESTAMP_PROVIDER_KEY = "@microsoft/live-share:TimestampProvider";

/**
 * Options used to configure the `LiveShareClient` class.
 */
export interface ILiveShareClientOptions {
    /**
     * Optional. Configuration to use when connecting to a custom Azure Fluid Relay instance.
     */
    readonly connection?:
        | AzureRemoteConnectionConfig
        | AzureLocalConnectionConfig;

    /**
     * Optional. A logger instance to receive diagnostic messages.
     */
    readonly logger?: ITelemetryBaseLogger;

    /**
     * Optional. Custom timestamp provider to use.
     */
    readonly timestampProvider?: ITimestampProvider;
}

/**
 * Client used to connect to fluid containers within a Microsoft Teams context.
 */
export class LiveShareClient {
    private static _host: ILiveShareHost = TestLiveShareHost.create(
        undefined,
        undefined
    );
    private static get host(): ILiveShareHost {
        if (typeof window === "undefined") {
            return this._host;
        }
        return ((window as any)[GLOBAL_LIVE_SHARE_HOST_WINDOW_KEY] || this._host) as ILiveShareHost;
    }
    private static set host(value: ILiveShareHost) {
        if (typeof window === "undefined") {
            this._host = value;
            return;
        }
        (window as any)[GLOBAL_LIVE_SHARE_HOST_WINDOW_KEY] = value;
    }
    private readonly _options: ILiveShareClientOptions;
    private static _timestampProvider: ITimestampProvider =
        new LocalTimestampProvider();
    private static get timestampProvider(): ITimestampProvider {
        if (typeof window === "undefined") {
            return this._timestampProvider;
        }
        return ((window as any)[GLOBAL_LIVE_SHARE_TIMESTAMP_PROVIDER_KEY] || this._timestampProvider) as ITimestampProvider;
    }
    private static set timestampProvider(value: ITimestampProvider) {
        if (typeof window === "undefined") {
            this._timestampProvider = value;
            return;
        }
        (window as any)[GLOBAL_LIVE_SHARE_TIMESTAMP_PROVIDER_KEY] = value;
    }
    
    private static get _roleVerifier(): IRoleVerifier {
        return new RoleVerifier(LiveShareClient.host);
    }

    /**
     * Creates a new `LiveShareClient` instance.
     * @param host Host for the current Live Share session.
     * @param options Optional. Configuration options for the client.
     */
    constructor(host: ILiveShareHost, options?: ILiveShareClientOptions) {
        // Validate host passed in
        if (!host || typeof host.getFluidTenantInfo != "function") {
            throw new Error(`LiveShareClient: host not passed in`);
        }

        // Save props
        LiveShareClient.host = new PolyfillHostDecorator(
            new LiveShareHostDecorator(host)
        );
        this._options = Object.assign({} as ILiveShareClientOptions, options);
    }

    /**
     * If true the client is configured to use a local test server.
     */
    public get isTesting(): boolean {
        return this._options.connection?.type == "local";
    }

    /**
     * Number of times the client should attempt to get the ID of the container to join for the
     * current context.
     */
    public maxContainerLookupTries = 3;

    /**
     * Connects to the fluid container for the current teams context.
     *
     * @remarks
     * The first client joining the container will create the container resulting in the
     * `onContainerFirstCreated` callback being called. This callback can be used to set the initial
     * state of of the containers object prior to the container being attached.
     * @param fluidContainerSchema Fluid objects to create.
     * @param onContainerFirstCreated Optional. Callback that's called when the container is first created.
     * @returns The fluid `container` and `services` objects to use along with a `created` flag that if true means the container had to be created.
     */
    public async joinContainer(
        fluidContainerSchema: ContainerSchema,
        onContainerFirstCreated?: (container: IFluidContainer) => void
    ): Promise<{
        container: IFluidContainer;
        services: AzureContainerServices;
        created: boolean;
    }> {
        performance.mark(`TeamsSync: join container`);
        try {
            // Configure role verifier and timestamp provider
            const pTimestampProvider = this.initializeTimestampProvider();

            // Initialize FRS connection config
            let config:
                | AzureRemoteConnectionConfig
                | AzureLocalConnectionConfig
                | undefined = this._options.connection;
            if (!config) {
                const frsTenantInfo =
                    await LiveShareClient.host.getFluidTenantInfo();

                // Compute endpoint
                let endpoint: string | undefined =
                    frsTenantInfo.serviceEndpoint;
                if (!endpoint) {
                    if (serviceEndpointMap.has(frsTenantInfo.serviceEndpoint)) {
                        endpoint = serviceEndpointMap.get(
                            frsTenantInfo.serviceEndpoint
                        );
                    } else {
                        throw new Error(
                            `LiveShareClient: Unable to find fluid endpoint for: ${frsTenantInfo.serviceEndpoint}`
                        );
                    }
                }

                // Is this a local config?
                if (frsTenantInfo.tenantId == "local") {
                    config = {
                        type: "local",
                        endpoint: endpoint!,
                        tokenProvider: new InsecureTokenProvider("", {
                            id: "123",
                            name: "Test User",
                        } as IUser),
                    };
                } else {
                    config = {
                        type: "remote",
                        tenantId: frsTenantInfo.tenantId,
                        endpoint: endpoint!,
                        tokenProvider: new LiveShareTokenProvider(
                            LiveShareClient.host
                        ),
                    } as AzureRemoteConnectionConfig;
                }
            }

            // Create FRS client
            const client = new AzureClient({
                connection: config,
                logger: this._options.logger,
            });

            // Create container on first access
            const pContainer = this.getOrCreateContainer(
                client,
                fluidContainerSchema,
                0,
                onContainerFirstCreated
            );

            // Wait in parallel for everything to finish initializing.
            const result = await Promise.all([
                pContainer,
                pTimestampProvider,
            ]);

            performance.mark(`TeamsSync: container connecting`);

            // Wait for containers socket to connect
            let connected = false;
            const { container, services } = result[0];
            container.on("connected", async () => {
                if (!connected) {
                    connected = true;
                    performance.measure(
                        `TeamsSync: container connected`,
                        `TeamsSync: container connecting`
                    );
                }

                // Register any new clientId's
                // - registerClientId() will only register a client on first use
                const connections =
                    services.audience.getMyself()?.connections ?? [];
                for (let i = 0; i < connections.length; i++) {
                    try {
                        const clientId = connections[i]?.id;
                        if (clientId) {
                            await LiveShareClient.host.registerClientId(
                                clientId
                            );
                        }
                    } catch (err: any) {
                        console.error(err.toString());
                    }
                }
            });

            return result[0];
        } finally {
            performance.measure(
                `TeamsSync: container joined`,
                `TeamsSync: join container`
            );
        }
    }

    /**
     * @hidden
     */
    protected async initializeTimestampProvider(): Promise<void> {
        if (!this.isTesting) {
            // Was a custom timestamp provider passed in.
            if (this._options.timestampProvider) {
                // Use configured one
                LiveShareClient.timestampProvider =
                    this._options.timestampProvider;
            } else {
                // Create a new host based timestamp provider
                LiveShareClient.timestampProvider = new HostTimestampProvider(
                    LiveShareClient.host
                );
            }

            // Start provider if needed
            if (
                typeof (LiveShareClient.timestampProvider as TimestampProvider)
                    .start == "function"
            ) {
                return (
                    LiveShareClient.timestampProvider as TimestampProvider
                ).start();
            }
        }

        return Promise.resolve();
    }

    private async getOrCreateContainer(
        client: AzureClient,
        fluidContainerSchema: ContainerSchema,
        tries: number,
        onInitializeContainer?: (container: IFluidContainer) => void
    ): Promise<{
        container: IFluidContainer;
        services: AzureContainerServices;
        created: boolean;
    }> {
        // Get container ID mapping
        const containerInfo = await LiveShareClient.host.getFluidContainerId();

        // Create container on first access
        if (containerInfo.shouldCreate) {
            return await this.createNewContainer(
                client,
                fluidContainerSchema,
                tries,
                onInitializeContainer
            );
        } else if (containerInfo.containerId) {
            return {
                created: false,
                ...(await client.getContainer(
                    containerInfo.containerId,
                    fluidContainerSchema
                )),
            };
        } else if (
            tries < this.maxContainerLookupTries &&
            containerInfo.retryAfter > 0
        ) {
            await this.wait(containerInfo.retryAfter);
            return await this.getOrCreateContainer(
                client,
                fluidContainerSchema,
                tries + 1,
                onInitializeContainer
            );
        } else {
            throw new Error(
                `LiveShareClient: timed out attempting to create or get container for current context.`
            );
        }
    }

    private async createNewContainer(
        client: AzureClient,
        fluidContainerSchema: ContainerSchema,
        tries: number,
        onInitializeContainer?: (container: IFluidContainer) => void
    ): Promise<{
        container: IFluidContainer;
        services: AzureContainerServices;
        created: boolean;
    }> {
        // Create and initialize container
        const { container, services } = await client.createContainer(
            fluidContainerSchema
        );
        if (onInitializeContainer) {
            onInitializeContainer(container);
        }

        // Attach container to service
        const newContainerId = await container.attach();

        // Attempt to save container ID mapping
        const containerInfo = await LiveShareClient.host.setFluidContainerId(
            newContainerId
        );
        if (containerInfo.containerState != ContainerState.added) {
            // Delete created container
            container.dispose();

            // Get mapped container ID
            return {
                created: false,
                ...(await client.getContainer(
                    containerInfo.containerId!,
                    fluidContainerSchema
                )),
            };
        } else {
            return { container, services, created: true };
        }
    }

    private wait(delay: number): Promise<void> {
        return new Promise((resolve) => {
            setTimeout(() => resolve(), delay);
        });
    }

    /**
     * Returns the current timestamp as the number of milliseconds sine the Unix Epoch.
     */
    public static getTimestamp(): number {
        return LiveShareClient.timestampProvider.getTimestamp();
    }

    /**
     * Verifies that a client has one of the specified roles.
     * @param clientId Client ID to inspect.
     * @param allowedRoles User roles that are allowed.
     * @returns True if the client has one of the specified roles.
     */
    public static verifyRolesAllowed(
        clientId: string,
        allowedRoles: UserMeetingRole[]
    ): Promise<boolean> {
        return LiveShareClient._roleVerifier.verifyRolesAllowed(
            clientId,
            allowedRoles
        );
    }

    public static getClientInfo(
        clientId: string
    ): Promise<IClientInfo | undefined> {
        return LiveShareClient.host.getClientInfo(clientId);
    }

    /**
     * Assigns a custom timestamp provider.
     * @param provider The timestamp provider to use.
     */
    public static setTimestampProvider(provider: ITimestampProvider): void {
        LiveShareClient.timestampProvider = provider;
    }
}
