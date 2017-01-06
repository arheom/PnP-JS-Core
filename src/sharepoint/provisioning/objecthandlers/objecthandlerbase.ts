"use strict";

import { HttpClient } from "../../../net/httpclient";
import { Logger } from "../../../utils/logging";

/**
 * Describes the Object Handler Base
 */
export class ObjectHandlerBase {
    public httpClient: HttpClient;
    private name: string;

    /**
     * Creates a new instance of the ObjectHandlerBase class
     */
    constructor(name: string) {
        this.name = name;
        this.httpClient = new HttpClient();
    }

    /**
     * Provisioning objects
     */
    public ProvisionObjects(objects: any, tokenParser: any, parameters?) {
        return new Promise((resolve, reject) => { resolve("Not implemented."); });
    }

    /**
     * Writes to Logger when scope has started
     */
    public scope_started() {
        Logger.write(`${this.name}: Code execution scope started`);
    }

    /**
     * Writes to Logger when scope has stopped
     */
    public scope_ended() {
        Logger.write(`${this.name}: Code execution scope stopped`);
    }
}
