import {
    DataObjectTypes,
    IDataObjectProps,
} from "@fluidframework/aqueduct";
import { IFluidLoadable } from "@fluidframework/core-interfaces";
import { LiveDataObject } from "./LiveDataObject";
import { LiveShareRuntime } from "./LiveShareRuntime";
import {
    ContainerSchema,
    LoadableObjectClass,
    LoadableObjectClassRecord,
} from "fluid-framework";

/**
 * Inject Live Share dependencies into your Fluid container schema.
 * This should only be done once, right before connecting to a container.
 * @remarks
 * Users should not use this method unless you are connecting to a container using `LiveShareClient`.
 * This is intended to be used when you are using another Fluid client, such as `AzureClient`.
 *
 * @param schema Fluid ContainerSchema you would like to inject the runtime into
 * @param liveRuntime LiveShareRuntime instance
 * @returns ContainerSchema with injected dependencies
 */
export function getLiveShareContainerSchemaProxy(
    schema: ContainerSchema,
    liveRuntime: LiveShareRuntime
): ContainerSchema {
    // Each container must proxy LiveDataObject classes separately.
    // This map is used to de-duplicate proxies for each class.
    const existingProxyRegistries = new Map<string, LoadableObjectClass<any>>();
    
    const initialObjectEntries = Object.entries(schema.initialObjects).map(
        ([key, ObjectClass]) => {
            return [key, getLiveDataObjectClassProxy(ObjectClass, liveRuntime, existingProxyRegistries)];
        }
    );
    const newInitialObjects: LoadableObjectClassRecord =
        Object.fromEntries(initialObjectEntries);

    return {
        initialObjects: newInitialObjects,
        dynamicObjectTypes: schema.dynamicObjectTypes?.map((ObjectClass) =>
            getLiveDataObjectClassProxy(ObjectClass, liveRuntime, existingProxyRegistries)
        ),
    };
}

/**
 * @hidden
 * Inject Live Share dependencies to relevant `LiveDataObject` derived classes.
 * Regular `DataObject` classes are not proxied.
 */
export function getLiveDataObjectClassProxy<TClass extends IFluidLoadable>(
    ObjectClass: LoadableObjectClass<any>,
    liveRuntime: LiveShareRuntime,
    existingProxyRegistries: Map<string, LoadableObjectClass<any>> = new Map(),
): LoadableObjectClass<TClass> {
    if (isLiveDataObject(ObjectClass)) {
        // We should only be proxying one Live Share DDS per type.
        // This is because Fluid attempts to de-duplicate by comparing classes, but we are dynamically creating proxies.
        // They then enforce this de-duplication using the factory type name, throwing an error in `parseDataObjectsFromSharedObjects`.
        // So, we ensure that we only create the proxy once per container.
        const typeName = (ObjectClass as any).TypeName;
        const CheckExisting = existingProxyRegistries.get(typeName);
        if (CheckExisting !== undefined) {
            return CheckExisting;
        }
        // Create a new proxy for this type and insert it into proxiedClasses
        const NewProxy = proxyLiveDataObjectClass(
            ObjectClass,
            liveRuntime
        ) as unknown as LoadableObjectClass<TClass>;
        existingProxyRegistries.set(typeName, NewProxy);
        return NewProxy;
    }
    return ObjectClass;
}

/**
 * @hidden
 */
function isLiveDataObject(value: any): value is typeof LiveDataObject {
    return value.LiveEnabled === true;
}

/**
 * @hidden
 * Create a new class extending LiveDataObject to inject in _liveRuntime
 */
function proxyLiveDataObjectClass<I extends DataObjectTypes = DataObjectTypes>(
    BaseClass: typeof LiveDataObject<I>,
    runtime: LiveShareRuntime
): typeof LiveDataObject<I> {
    class ProxiedBaseClass extends (BaseClass as unknown as new (
        props: IDataObjectProps<I>
    ) => LiveDataObject<I>) {
        constructor(props: IDataObjectProps<I>) {
            super(props);
            this["liveRuntime"] = runtime;
        }
    }

    const DynamicClass: typeof LiveDataObject<I> = class extends BaseClass {
        public static TypeName = (BaseClass as any).TypeName;
        public static readonly factory = new Proxy((BaseClass as any).factory, {
            get: function (target, prop, receiver) {
                if (prop === "ctor") {
                    return ProxiedBaseClass;
                }
                return Reflect.get(target, prop, receiver);
            },
        });
    };

    return DynamicClass;
}