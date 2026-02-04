// Type overrides for reactflow to fix d3-dispatch compatibility issues
declare module 'reactflow' {
    export * from 'reactflow/dist/esm';
}

// Override problematic d3-dispatch types
declare module 'd3-dispatch' {
    export function dispatch<T extends string>(...types: T[]): any;
}
