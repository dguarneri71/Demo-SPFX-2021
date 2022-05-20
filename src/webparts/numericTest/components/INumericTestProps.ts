import { WebPartContext } from '@microsoft/sp-webpart-base';
import { TestItem } from '../../../classes/commons/TestItem';

export interface INumericTestProps {
  description: string;
  wpContext: WebPartContext;
  libraryId?: string;
  siteUrl?: string;
}

export interface INumericTestState {
  items: TestItem[];
}