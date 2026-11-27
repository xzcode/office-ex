// SPDX-License-Identifier: MIT
import {makeImageResource,createDocument,DocumentSession} from '@office-ex/core';
import {insertImageAfter,replaceParagraphRange,insertPageBreakAfter} from '@office-ex/writer';
export const PNG=Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAgAAAAGCAIAAABxZ0isAAAAHklEQVR4nGP8z8Cgl7mKAQMw6mITZWBgYMIqSl0JAOavA4uzO/5KAAAAAElFTkSuQmCC','base64');
export const JPEG=Buffer.from('/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/wAARCAAGAAgDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDnQMDAooor6Y8Btt3Z/9k=','base64');
export function imageFixture(){
  const s=new DocumentSession(createDocument('writer','image-fixture','Pictures'),{saved:true});
  s.execute(replaceParagraphRange(s.snapshot(),'paragraph-1',0,0,'Picture example','heading'));
  s.execute(insertImageAfter(s.snapshot(),'paragraph-1',makeImageResource(PNG),'png'));
  s.execute(insertImageAfter(s.snapshot(),'png-p',makeImageResource(JPEG),'jpeg'));
  return s;
}
