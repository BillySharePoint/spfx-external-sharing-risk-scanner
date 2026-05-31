import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Icon } from '@fluentui/react/lib/Icon';

export interface IEmptyStateProps {
    message?: string;
    iconName?: string;
}

export const EmptyState: React.FC<IEmptyStateProps> = ({
    message = 'No external users detected — this is a good sign!',
    iconName = 'CheckMark',
}) => {
    return (
        <Stack
            horizontalAlign="center"
            verticalAlign="center"
            tokens={{ childrenGap: 12, padding: 24 }}
            style={{
                backgroundColor: '#f3f9f4',
                borderRadius: 4,
                border: '1px solid #dff6dd',
            }}
        >
            <Icon
                iconName={iconName}
                style={{ fontSize: 32, color: '#107c10' }}
                aria-hidden="true"
            />
            <Text
                variant="medium"
                style={{ color: '#107c10', textAlign: 'center' }}
            >
                {message}
            </Text>
        </Stack>
    );
};
