import * as React from 'react';
import { useState, useEffect, useRef } from "react";
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';

import { useId, useBoolean } from '@fluentui/react-hooks';
import IConfirmDialogProps from './IConfirmDialogProps';

const dialogStyles = { main: { maxWidth: 450 } };

const dragOptions = {
    moveMenuItemText: 'Move',
    closeMenuItemText: 'Close',
    keepInBounds: true,
};

export default function ConfirmDialog(props: IConfirmDialogProps) : JSX.Element {
    const [isDraggable, { toggle: toggleIsDraggable }] = useBoolean(false);
    const labelId: string = useId('dialogLabel');
    const subTextId: string = useId('subTextLabel');


    const dialogContentProps = {
        type: DialogType.normal,
        title: props.title,
        closeButtonAriaLabel: 'Close',
        subText: props.message
    };

    const modalProps = React.useMemo(
        () => ({
          titleAriaId: labelId,
          subtitleAriaId: subTextId,
          isBlocking: false,
          styles: dialogStyles,
          dragOptions: isDraggable ? dragOptions : undefined,
        }),
        [isDraggable, labelId, subTextId],
    );

    return <Dialog hidden={ !props.show } dialogContentProps={ dialogContentProps }> 
                <DialogFooter>
                    <PrimaryButton onClick={ () => props.onClick(true) } text="Yes" />
                    <DefaultButton onClick={ ()=> props.onClick(false) } text="Cancel" />
                </DialogFooter>
            </Dialog>;
}