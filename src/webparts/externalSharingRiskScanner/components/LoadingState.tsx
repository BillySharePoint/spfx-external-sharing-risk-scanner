import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';

export interface ILoadingStateProps {
    message?: string;
}

export const LoadingState: React.FC<ILoadingStateProps> = ({
    message = 'Scanning site for external sharing risk signals...',
}) => {
    return (
        <Stack
            horizontalAlign="center"
            verticalAlign="center"
            tokens={{ childrenGap: 16, padding: 40 }}
            aria-live="polite"
            aria-busy="true"
        >
            <Spinner
                size={SpinnerSize.large}
                label={message}
                ariaLive="assertive"
            />
            <Text variant="small" style={{ color: '#605e5c', marginTop: 8 }}>
                This may take a few moments depending on site size and permissions.
            </Text>
        </Stack>
    );
};
