// SPDX-License-Identifier: MIT
// The emitted worker retains Mozilla PDF.js's Apache-2.0 license. This small
// original adapter exposes PDF ActualText, which upstream 6.5.438 otherwise
// omits from marked-content extraction. Rendering and parsing are unchanged.
import { createHash } from 'node:crypto';
export const PDF_WORKER_UPSTREAM_SHA256 = '43710009ea26bfbcb615524946e46e3aaf8ad927514f5e465e8ff6d18044ffe9';
export function preparePdfWorker(source) {
  if (createHash('sha256').update(source).digest('hex') !== PDF_WORKER_UPSTREAM_SHA256) throw new Error('PDF.js worker changed; review ActualText adapter before upgrading');
  const anchor = '                type: "beginMarkedContentProps",\n                id: Number.isInteger(mcid) ? `${self.idFactory.getPageObjId()}_mc${mcid}` : null,\n                tag: args[0] instanceof Name ? args[0].name : null';
  if (source.split(anchor).length !== 2) throw new Error('PDF.js marked-content adapter anchor is not unique');
  const extra = `,
                // Office EX ActualText bridge (Apache-2.0 worker modification).
                actualText: (() => {
                  const properties = args[1] instanceof Dict ? args[1] : args[1] instanceof Name ? resources.get("Properties")?.get(args[1].name) : null;
                  const text = properties instanceof Dict ? properties.get("ActualText") : null;
                  if (typeof text !== "string") return null;
                  if (text.length > 4000000) throw new Error("PDF ActualText budget exceeded");
                  return stringToPDFString(text);
                })()`;
  return source.replace(anchor, anchor + extra);
}
