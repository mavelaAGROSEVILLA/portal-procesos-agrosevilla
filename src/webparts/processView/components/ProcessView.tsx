import * as React from 'react';
import type { IProcessViewProps } from './IProcessViewProps';
import Home from '../../../pages/Home/Home';

const ProcessView: React.FunctionComponent<IProcessViewProps> = ({ sp }) => <Home sp={sp} />;

export default ProcessView;
