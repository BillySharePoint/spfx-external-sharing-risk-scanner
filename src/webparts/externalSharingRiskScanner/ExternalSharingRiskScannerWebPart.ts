import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import { SPFI, spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/site-users/web';
import '@pnp/sp/site-groups/web';
import '@pnp/sp/lists';
import '@pnp/sp/security/list';
import '@pnp/sp/security/web';

import {
    ExternalSharingRiskScanner,
    IExternalSharingRiskScannerProps,
} from './components/ExternalSharingRiskScanner';
import { LoggerService, LogLevel } from './services/LoggerService';

import * as strings from 'ExternalSharingRiskScannerWebPartStrings';

export interface IExternalSharingRiskScannerWebPartProps {
    internalDomains: string;
}

export default class ExternalSharingRiskScannerWebPart extends BaseClientSideWebPart<IExternalSharingRiskScannerWebPartProps> {
    private _sp!: SPFI;

    protected async onInit(): Promise<void> {
        await super.onInit();

        // Initialize PnPjs with SPFx context
        this._sp = spfi().using(SPFx(this.context));

        // Set logger level based on environment
        if (DEBUG) {
            LoggerService.setLevel(LogLevel.Debug);
            LoggerService.debug('WebPart', 'Debug mode enabled');
        } else {
            LoggerService.setLevel(LogLevel.Warning);
        }

        LoggerService.info('WebPart', `Initialized for site: ${this.context.pageContext.web.absoluteUrl}`);
    }

    public render(): void {
        const element: React.ReactElement<IExternalSharingRiskScannerProps> = React.createElement(
            ExternalSharingRiskScanner,
            {
                sp: this._sp,
                internalDomains: this.properties.internalDomains || '',
                siteUrl: this.context.pageContext.web.absoluteUrl,
            }
        );

        ReactDom.render(element, this.domElement);
    }

    protected onDispose(): void {
        ReactDom.unmountComponentAtNode(this.domElement);
    }

    protected get dataVersion(): Version {
        return Version.parse('1.0');
    }

    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
        return {
            pages: [
                {
                    header: {
                        description: strings.PropertyPaneDescription,
                    },
                    groups: [
                        {
                            groupName: strings.BasicGroupName,
                            groupFields: [
                                PropertyPaneTextField('internalDomains', {
                                    label: strings.InternalDomainsFieldLabel,
                                    description: strings.InternalDomainsDescription,
                                    multiline: true,
                                    rows: 3,
                                    placeholder: 'contoso.com, contoso.onmicrosoft.com',
                                }),
                            ],
                        },
                    ],
                },
            ],
        };
    }
}
