import * as React from "react";
import { Spinner, SpinnerSize } from "@fluentui/react";
import { Icon } from '@fluentui/react/lib/Icon';

const MyIcon = () => <Icon iconName="CompassNW" />;


export interface ButtonProps {
}

export default class Button extends React.Component<ButtonProps> {
  render() {
    const onClick = async () => {
        return Word.run(async (context) => {
            const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);
            paragraph.font.color = "blue";
            await context.sync();
        });     
    }
    return (
        <button onClick={onClick}>
            <MyIcon />
            click me
        </button>
    );
  }
}
