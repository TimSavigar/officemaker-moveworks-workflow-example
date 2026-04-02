import { createDocument, printCreateSummary } from '../src/officemaker-client.mjs';

const documentObject = {
  type: 'document',
  content: {
    children: [
      { type: 'paragraph', children: [{ type: 'text', text: 'Dear Employee,' }] },
      { type: 'paragraph', children: [{ type: 'text', text: 'This document was created from the OfficeMaker Moveworks workflow starter.' }] },
      { type: 'paragraph', children: [{ type: 'text', text: 'Kind regards,' }] },
      { type: 'paragraph', children: [{ type: 'text', text: 'Moveworks Agent' }] }
    ]
  }
};

const result = await createDocument({ documentType: 'word', fileName: 'moveworks-letter', documentObject });
printCreateSummary(result);
