import {TokenDefinition} from "./tokendefinition";
import {TokensUtils} from "./tokensutils";

export class GroupIdToken extends TokenDefinition {
    private _groupId: number = 0;

    constructor(web: SP.Web, name: string, groupId: number) {
        super(web, ["<<groupid:" + TokensUtils.escapeRegExp(name) + ">>"]);
        this._groupId = groupId;
    }

    public getReplaceValue(): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            if (this.cacheValue == null || this.cacheValue === "") {
                this.cacheValue = this._groupId.toString();
            }
            resolve(this.cacheValue);
        });
    }
}
