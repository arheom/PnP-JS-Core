export class TokenDefinition {
    public web: SP.Web;

    protected cacheValue: string;

    private _tokens: string[];

    constructor(web: SP.Web, token: string[]) {
        this._tokens = token;
        this.web = web;
    }

    public getTokens(): string[] {
        return this._tokens;
    }

    public getRegex(): RegExp[] {
        let regexs = new RegExp[this._tokens.length];
        for (let q = 0; q < this._tokens.length; q++) {
            regexs[q] = new RegExp(this._tokens[q], "i");
        }
        return regexs;
    }

    public getRegexForToken(token: string): RegExp {
        return new RegExp(token, "i");
    }

    public getTokenLength(): number {
        return Math.max.apply(Math, this._tokens.map((item) => { return item.length; }));
    }

    public getReplaceValue(): Promise<string> {
        return null;
    }

    public clearCache(): void {
        this.cacheValue = null;
    }
}
