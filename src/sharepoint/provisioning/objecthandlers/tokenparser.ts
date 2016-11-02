import {TokenDefinition} from "./tokendefinitions/tokendefinition";
import {SiteCollectionTermStoreIdToken} from "./tokendefinitions/sitecollectiontermstoreidtoken";
import {TermSetIdToken} from "./tokendefinitions/termsetidtoken";
import {TermStoreIdToken} from "./tokendefinitions/termstoreidtoken";
import {ListIdToken} from "./tokendefinitions/listidtoken";
import {ListUrlToken} from "./tokendefinitions/listurltoken";
import {ContentTypeIdToken} from "./tokendefinitions/contenttypeidtoken";
import {ParameterToken} from "./tokendefinitions/parametertoken";
import {FieldTitleToken} from "./tokendefinitions/fieldtitletoken";
import {RoleDefinitionToken} from "./tokendefinitions/roledefinitiontoken";
import {GroupIdToken} from "./tokendefinitions/groupidtoken";

export class TokenParser {
    public _web: SP.Web;

    private _tokens: TokenDefinition[] = [];

    public getTokens(): TokenDefinition[] {
        return this._tokens;
    }


    public addToken(tokenDefinition: TokenDefinition): void {
        this._tokens.push(tokenDefinition);
        // ORDER IS IMPORTANT!
        let sortedTokens = this._tokens.sort((a, b) => { return (a.getTokenLength() - b.getTokenLength()); });
        this._tokens = sortedTokens;
    }

    public tokenParser(web: SP.Web, template: any): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            let context = web.get_context();
            let rootWeb = (context as SP.ClientContext).get_site().get_rootWeb();
            context.load(web, "ServerRelativeUrl", "Language");
            context.load(rootWeb, "ServerRelativeUrl");
            context.executeQueryAsync(() => {
                this._web = web;
                this._tokens = [];
                this._tokens.push(new SiteCollectionTermStoreIdToken(web));
                // Add lists
                let lists = this._web.get_lists();
                context.load(lists, "Include(Title, RootFolder, RootFolder.ServerRelativeUrl)");
                context.executeQueryAsync(() => {
                    lists.get_data().forEach((list) => {
                        this._tokens.push(new ListIdToken(web, list.get_title(), list.get_id()));
                        this._tokens.push(
                            new ListUrlToken(web,
                                list.get_title(),
                                list.get_rootFolder().get_serverRelativeUrl().substring(this._web.get_serverRelativeUrl().length + 1)));
                    });
                },
                    (sender, error) => { reject(error); });
                if (web.get_id().toString() !== rootWeb.get_id().toString()) {
                    // sub site
                    let sLists = rootWeb.get_lists();
                    context.load(sLists, "Include(Title, RootFolder, RootFolder.ServerRelativeUrl)");
                    context.executeQueryAsync(() => {
                        sLists.get_data().forEach((list) => {
                            this._tokens.push(new ListIdToken(web, list.get_title(), list.get_id()));
                            this._tokens.push(
                                new ListUrlToken(web,
                                    list.get_title(),
                                    list.get_rootFolder().get_serverRelativeUrl().substring(this._web.get_serverRelativeUrl().length + 1)));
                        });
                    },
                        (sender, error) => { reject(error); });
                }

                // Add content types
                let contentTypes = web.get_contentTypes();
                context.load(contentTypes, "Include(StringId, Name)");
                context.executeQueryAsync(() => {
                    contentTypes.get_data().forEach((ct) => {
                        this._tokens.push(new ContentTypeIdToken(web, ct.get_name(), ct.get_stringId()));
                    });
                },
                    (sender, error) => { reject(error); });
                // Add parameters
                if (template.hasOwnProperty("Parameters") && template.Parameters.length > 0) {
                    template.Parameters.forEach((param) => {
                        this._tokens.push(new ParameterToken(web, param.Key, param.Value));
                    });
                }
                // Add TermSetIds
                let session = SP.Taxonomy.TaxonomySession.getTaxonomySession(context as SP.ClientContext);
                let termStores = session.get_termStores();
                context.load(termStores);
                context.executeQueryAsync(() => {
                    termStores.get_data().forEach((ts) => {
                        this._tokens.push(new TermStoreIdToken(web, ts.get_name(), ts.get_id()));
                    });
                }, (sender, error) => { reject(error); });

                let termStore = session.getDefaultSiteCollectionTermStore();
                let termGroups = termStore.get_groups();
                context.load(termStore);
                context.load(termGroups, "Include(Name, TermSets.Include(Name))");
                context.executeQueryAsync(() => {
                    termGroups.get_data().forEach((termGroup) => {
                        termGroup.get_termSets().get_data().forEach((termSet) => {
                            this._tokens.push(new TermSetIdToken(web, termGroup.get_name(), termSet.get_name(), termSet.get_id()));
                        });
                    });
                }, (sender, error) => { reject(error); });

                let fields = web.get_fields();
                context.load(fields, "Include(Title, InternalName)");
                context.executeQueryAsync(() => {
                    fields.get_data().forEach((field) => {
                        this._tokens.push(new FieldTitleToken(web, field.get_internalName(), field.get_title()));
                    });
                }, (sender, error) => { reject(error); });
                if (web.get_id().toString() !== rootWeb.get_id().toString()) {
                    // sub site
                    let sFields = rootWeb.get_fields();
                    context.load(sFields, "Include(Title, InternalName)");
                    context.executeQueryAsync(() => {
                        sFields.get_data().forEach((field) => {
                            this._tokens.push(new FieldTitleToken(rootWeb, field.get_internalName(), field.get_title()));
                        });
                    }, (sender, error) => { reject(error); });
                }
                // TODO: Handle resources

                // Add Role Definitions
                let roleDefinitions = web.get_roleDefinitions();
                context.load(roleDefinitions, "Include(RoleTypeKind)");
                context.executeQueryAsync(() => {
                    roleDefinitions.get_data().filter((roleDef) => roleDef.get_roleTypeKind() !== SP.RoleType.none).forEach((roleDef) => {
                        this._tokens.push(new RoleDefinitionToken(web, roleDef));
                    });
                }, (sender, error) => { reject(error); });

                // Add Groups

                let groups = web.get_siteGroups();
                context.load(groups, "Include(Title)");
                context.executeQueryAsync(() => {
                    groups.get_data().forEach((group) => {
                        this._tokens.push(new GroupIdToken(web, group.get_title(), group.get_id()));
                    });
                }, (sender, error) => { reject(error); });
                context.load(web, "AssociatedVisitorGroup", "AssociatedMemberGroup", "AssociatedOwnerGroup");
                context.executeQueryAsync(() => {
                    if (!web.get_associatedVisitorGroup().get_serverObjectIsNull()) {
                        this._tokens.push(new GroupIdToken(web, "associatedvisitorgroup", web.get_associatedVisitorGroup().get_id()));
                    }
                    if (!web.get_associatedMemberGroup().get_serverObjectIsNull()) {
                        this._tokens.push(new GroupIdToken(web, "associatedmembergroup", web.get_associatedMemberGroup().get_id()));
                    }
                    if (!web.get_associatedOwnerGroup().get_serverObjectIsNull()) {
                        this._tokens.push(new GroupIdToken(web, "associatedownergroup", web.get_associatedOwnerGroup().get_id()));
                    }
                }, (sender, error) => { reject(error); });

                let sortedTokens = this._tokens.sort((a, b) => { return (a.getTokenLength() - b.getTokenLength()); });
                this._tokens = sortedTokens;
            }, (sender, error) => { reject(error); });
        });
    }

    public rebase(web: SP.Web): Promise<void> {
        return new Promise<void>((resolve, reject) => {
            web.get_context().load(web, "ServerRelativeUrl", "Language");
            web.get_context().executeQueryAsync(() => {
                this._web = web;
                this._tokens.forEach((token) => {
                    token.clearCache();
                    token.web = web;
                });
                resolve();
            }, (sender, error) => { reject(error); });
        });
    }

    public parseString(input: string): Promise<string> {
        return this.parseStringWithSkip(input, null);
    }

    public getLeftOverTokens(input: string): string[] {
        let values: string[] = [];
        let matches = input.match(new RegExp("(?<guid>\<\S{8}-\S{4}-\S{4}-\S{4}-\S{12}?\>)|(?<token>\<.+?\>)"));
        matches.forEach(match => {
            values.push(match);
        });
        return values;
    }

    public parseStringWithSkip(input: string, tokensToSkip: string[]): Promise<string> {
        return new Promise<string>((resolve, reject) => {
            let origInput = input;
            if (input != null && input !== "") {
                this._tokens.forEach(token => {
                    token.getReplaceValue().then(tokenValue => {
                        if (tokensToSkip != null) {
                            let filteredTokens = token.getTokens().filter((filteredToken) => { return tokensToSkip.indexOf(filteredToken) < 0; });
                            if (filteredTokens.length > 0) {
                                filteredTokens.forEach((filteredToken) => {
                                    let regex = token.getRegexForToken(filteredToken);
                                    if (input.match(regex)) {
                                        input = input.replace(regex, tokenValue);
                                    }
                                });
                            }
                        } else {
                            let matchingTokens: RegExp[] = token.getRegex().filter(regex => input.match(regex).length > 0);
                            matchingTokens.forEach(regex => {
                                input = input.replace(regex, tokenValue);
                            });
                        }
                        resolve(input);
                    });
                });
            }

            while (origInput !== input) {
                this._tokens.forEach(token => {
                    token.getReplaceValue().then(tokenValue => {
                        origInput = input;
                        if (tokensToSkip != null) {
                            let filteredTokens = token.getTokens().filter(filteredToken => tokensToSkip.indexOf(filteredToken) > 0);
                            if (filteredTokens.length > 0) {
                                filteredTokens.forEach(filteredToken => {
                                    let regex = token.getRegexForToken(filteredToken);
                                    if (input.match(regex)) {
                                        input = input.replace(regex, tokenValue);
                                    }
                                });
                            } else {
                                let filteredRegexs = token.getRegex().filter(regex => input.match(regex).length > 0);
                                filteredRegexs.forEach(regex => {
                                    origInput = input;
                                    input = input.replace(regex, tokenValue);
                                });
                            }
                        }
                        resolve(input);
                    });
                });
            }
        });
    }

}
