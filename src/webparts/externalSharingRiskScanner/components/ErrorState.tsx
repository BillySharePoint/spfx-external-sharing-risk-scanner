import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Link } from '@fluentui/react/lib/Link';
import { IScannerError } from '../models/IScannerError';
import { getFriendlyErrorMessage } from '../utils/errorUtils';

export interface IErrorStateProps {
    error: IScannerError;
    showTechnicalDetails?: boolean;
}

export const ErrorState: React.FC<IErrorStateProps> = ({
    error,
    showTechnicalDetails = false,
}) => {
    const [showDetails, setShowDetails] = React.useState(false);
    const friendlyMessage = getFriendlyErrorMessage(error);

    return (
        <Stack tokens={{ childrenGap: 8 }} aria-live="polite">
            <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline={true}
                aria-label={`Error from ${error.source}: ${friendlyMessage}`}
            >
                <Stack tokens={{ childrenGap: 4 }}>
                    <Text variant="mediumPlus" style={{ fontWeight: 600 }}>
                        {error.source}
                    </Text>
                    <Text>{friendlyMessage}</Text>
                    {showTechnicalDetails && error.technicalDetails && (
                        <Stack tokens={{ childrenGap: 4 }}>
                            <Link onClick={() => setShowDetails(!showDetails)}>
                                {showDetails ? 'Hide technical details' : 'Show technical details'}
                            </Link>
                            {showDetails && (
                                <Text
                                    variant="small"
                                    style={{
                                        fontFamily: 'monospace',
                                        whiteSpace: 'pre-wrap',
                                        backgroundColor: '#faf9f8',
                                        padding: 8,
                                        borderRadius: 2,
                                        maxHeight: 200,
                                        overflow: 'auto',
                                    }}
                                >
                                    {error.technicalDetails}
                                </Text>
                            )}
                        </Stack>
                    )}
                </Stack>
            </MessageBar>
        </Stack>
    );
};
