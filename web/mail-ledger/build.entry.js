import MsgReader from "@kenjiuno/msgreader";
import { decompressRTF } from "@kenjiuno/decompressrtf";

const g = typeof window !== "undefined" ? window : globalThis;
g.MsgReader = MsgReader;
g.decompressRTF = decompressRTF;
