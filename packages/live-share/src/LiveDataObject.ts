import {
    DataObject,
    DataObjectTypes,
    IDataObjectProps,
} from "@fluidframework/aqueduct";
import { LiveShareRuntime } from "./LiveShareRuntime";
import { assert } from "@fluidframework/common-utils";
import { UserMeetingRole } from "./interfaces";
import { waitUntilConnected } from "./internals";

/**
 * Extends Fluid's DataObject class. Intended for use with Live Share custom DDS's that rely on a `ILiveShareHost`.
 */
export abstract class LiveDataObject<
    I extends DataObjectTypes = DataObjectTypes
> extends DataObject<I> {
    /**
     * @hidden
     */
    public static LiveEnabled = true;

    /**
     * @hidden
     */
    protected _allowedRoles: UserMeetingRole[] = [];

    /**
     * @hidden
     */
    private _liveRuntime: LiveShareRuntime | null = null;

    /**
     * @internal
     * `LiveShareRuntime` instance
     * @remarks
     * You should usually not set this value to a DDS after calling `.initialize()`, but there is nothing preventing it.
     */
    protected get liveRuntime(): LiveShareRuntime {
        assert(
            this._liveRuntime !== null,
            "LiveShareRuntime not initialized. Ensure your Fluid `ContainerSchema` was first wrapped inside of `getLiveShareSchema`, or use `.joinContainer()` in `LiveShareClient`."
        );
        return this._liveRuntime;
    }

    public constructor(props: IDataObjectProps<I>) {
        super(props);
    }

    /**
     * Waits until connected and gets the most recent clientId
     * @returns clientId
     */
    protected waitUntilConnected(): Promise<string> {
        return waitUntilConnected(this.runtime);
    }

    /**
     * Verify that the user has the required roles
     * @returns boolean true if user has the required roles
     */
    protected async verifyLocalUserRoles(): Promise<boolean> {
        const clientId = await this.waitUntilConnected();
        return this.liveRuntime.verifyRolesAllowed(
            clientId,
            this._allowedRoles ?? []
        );
    }

    /**
     * @hidden
     * Dependency injection setter for `LiveShareRuntime`.
     */
    protected __dangerouslySetLiveRuntime(value: LiveShareRuntime) {
        this._liveRuntime = value;
    }

    /**
     * @hidden
     * Utility function that lets you run a function if successful, or throw a consistent Error if not.
     * @param fnSuccess function to run if user has needed roles
     * @returns TResponse if successful. Should never throw an error.
     */
    public async onLocalUserAllowed(fnSuccess: () => void): Promise<void> {
        const valid = await this.verifyLocalUserRoles();
        if (!valid) return;
        fnSuccess();
    }
}
