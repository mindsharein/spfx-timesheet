import * as React from 'react';
import {
    MessageBar,
    MessageBarType,
} from '@fluentui/react';

export interface IMessageProps {
    text: string;
    type: MessageType
    reset(): void;
}

export enum MessageType {
    success,
    error
}

export default function Message(p: IMessageProps) : JSX.Element {
    return (
        <MessageBar
            messageBarType={ p.type == MessageType.success ? MessageBarType.success : MessageBarType.error }
            isMultiline={false}
            onDismiss={ p.reset }
            dismissButtonAriaLabel="Close"
        >
      { p.text }

    </MessageBar>
    );
};