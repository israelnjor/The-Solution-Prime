/**
 * Password-protected documents, as Word protects them (File → Info → Protect Document → Encrypt with
 * Password): ECMA-376 Agile Encryption (MS-OFFCRYPTO) — the .docx package encrypted with AES-256 (CBC)
 * under a key derived from the password with SHA-512 (100,000 rounds), in a compound file beside the
 * encryption's description, with an HMAC that shows tampering. WebCrypto does the cryptography.
 */
import { isCfb, readCfb, writeCfb } from "./cfb.js";

const subtle = () => globalThis.crypto.subtle;
const cat = (...parts: Uint8Array[]) => { const out = new Uint8Array(parts.reduce((a, p) => a + p.length, 0)); let o = 0; for (const p of parts) { out.set(p, o); o += p.length; } return out; };
const le32 = (n: number) => { const b = new Uint8Array(4); new DataView(b.buffer).setUint32(0, n, true); return b; };
const sha512 = async (data: Uint8Array) => new Uint8Array(await subtle().digest("SHA-512", data as Uint8Array<ArrayBuffer>));
const random = (n: number) => globalThis.crypto.getRandomValues(new Uint8Array(n));
const b64 = (b: Uint8Array) => { let s = ""; for (const x of b) s += String.fromCharCode(x); return btoa(s); };
const unb64 = (s: string) => Uint8Array.from(atob(s), (c) => c.charCodeAt(0));

const BLOCK = {
  verifierInput: new Uint8Array([0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79]),
  verifierHash: new Uint8Array([0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e]),
  keyValue: new Uint8Array([0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6]),
  integrityKey: new Uint8Array([0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6]),
  integrityValue: new Uint8Array([0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33]),
};

/** AES-CBC without padding (WebCrypto pads: the padding block is dropped, or made up for decrypting). */
async function aes(encrypt: boolean, key: Uint8Array, iv: Uint8Array, data: Uint8Array): Promise<Uint8Array> {
  const k = await subtle().importKey("raw", key as Uint8Array<ArrayBuffer>, "AES-CBC", false, [encrypt ? "encrypt" : "decrypt"]);
  if (encrypt) {
    const out = new Uint8Array(await subtle().encrypt({ name: "AES-CBC", iv: iv as Uint8Array<ArrayBuffer> }, k, data as Uint8Array<ArrayBuffer>));
    return out.subarray(0, data.length);
  }
  // A last block that decrypts to a full block of padding: the encryption of nothing, chained on the last block.
  const last = data.subarray(data.length - 16);
  const ek = await subtle().importKey("raw", key as Uint8Array<ArrayBuffer>, "AES-CBC", false, ["encrypt"]);
  const pad = new Uint8Array(await subtle().encrypt({ name: "AES-CBC", iv: last as Uint8Array<ArrayBuffer> }, ek, new Uint8Array(0)));
  return new Uint8Array(await subtle().decrypt({ name: "AES-CBC", iv: iv as Uint8Array<ArrayBuffer> }, k, cat(data, pad) as Uint8Array<ArrayBuffer>));
}

/** The password's hash after the spin count's rounds (H_n). */
async function passwordHash(password: string, salt: Uint8Array, spins: number): Promise<Uint8Array> {
  const utf16 = new Uint8Array(password.length * 2);
  for (let i = 0; i < password.length; i++) { const c = password.charCodeAt(i); utf16[i * 2] = c & 0xff; utf16[i * 2 + 1] = c >> 8; }
  let h = await sha512(cat(salt, utf16));
  const buf = new Uint8Array(4 + 64);
  const dv = new DataView(buf.buffer);
  for (let i = 0; i < spins; i++) { dv.setUint32(0, i, true); buf.set(h, 4); h = await sha512(buf); }
  return h;
}
const keyFor = async (h: Uint8Array, block: Uint8Array, bytes: number) => {
  const k = await sha512(cat(h, block));
  if (k.length >= bytes) return k.subarray(0, bytes);
  const out = new Uint8Array(bytes).fill(0x36);
  out.set(k);
  return out;
};
const ivFor = async (salt: Uint8Array, block: Uint8Array, size: number) => (await sha512(cat(salt, block))).subarray(0, size);

/** A password-protected Word document (an encrypted package in a compound file). */
export function isEncrypted(bytes: Uint8Array): boolean {
  if (!isCfb(bytes)) return false;
  try { const s = readCfb(bytes); return s.has("EncryptionInfo") && s.has("EncryptedPackage"); } catch { return false; }
}

export class WrongPasswordError extends Error {
  constructor() { super("The password is not correct."); }
}

/** The .docx inside a password-protected document. Throws WrongPasswordError for a wrong password. */
export async function decryptDocx(bytes: Uint8Array, password: string): Promise<Uint8Array> {
  const streams = readCfb(bytes);
  const info = streams.get("EncryptionInfo");
  const pkg = streams.get("EncryptedPackage");
  if (!info || !pkg) throw new Error("This file is not a password-protected Word document.");
  const idv = new DataView(info.buffer, info.byteOffset, info.byteLength);
  if (idv.getUint16(0, true) !== 4 || idv.getUint16(2, true) !== 4) throw new Error("This document uses an older kind of password protection; open it in Word and save it again to protect it with a current one.");
  const xml = new TextDecoder().decode(info.subarray(8));
  const attr = (el: string, name: string) => new RegExp(`<(?:\\w+:)?${el}\\b[^>]*\\s${name}="([^"]*)"`).exec(xml)?.[1];
  if (attr("keyData", "cipherAlgorithm") !== "AES" || !/^SHA-?512$/i.test(attr("encryptedKey", "hashAlgorithm") ?? "") || !/^SHA-?512$/i.test(attr("keyData", "hashAlgorithm") ?? "")) {
    throw new Error("This document's encryption (not SHA-512) is not supported.");
  }
  const keyBits = Number(attr("encryptedKey", "keyBits"));
  const spins = Number(attr("encryptedKey", "spinCount"));
  const pSalt = unb64(attr("encryptedKey", "saltValue")!);
  const h = await passwordHash(password, pSalt, spins);
  // The verifier: the password is right when the hash of the decrypted input matches.
  const input = await aes(false, await keyFor(h, BLOCK.verifierInput, keyBits / 8), pSalt, unb64(attr("encryptedKey", "encryptedVerifierHashInput")!));
  const expected = await aes(false, await keyFor(h, BLOCK.verifierHash, keyBits / 8), pSalt, unb64(attr("encryptedKey", "encryptedVerifierHashValue")!));
  const actual = await sha512(input.subarray(0, Number(attr("encryptedKey", "saltSize") ?? 16)));
  if (!actual.every((b, i) => b === expected[i])) throw new WrongPasswordError();
  const secret = (await aes(false, await keyFor(h, BLOCK.keyValue, keyBits / 8), pSalt, unb64(attr("encryptedKey", "encryptedKeyValue")!))).subarray(0, Number(attr("keyData", "keyBits")) / 8);
  const kSalt = unb64(attr("keyData", "saltValue")!);
  const blockSize = Number(attr("keyData", "blockSize") ?? 16);
  const size = Number(new DataView(pkg.buffer, pkg.byteOffset, 8).getBigUint64(0, true));
  const out = new Uint8Array(Math.ceil(size / 4096) * 4096 + 16);
  for (let i = 0, o = 8; o < pkg.length; i++, o += 4096) {
    const seg = pkg.subarray(o, Math.min(pkg.length, o + 4096));
    const len = seg.length - (seg.length % blockSize);
    if (!len) break;
    out.set(await aes(false, secret, await ivFor(kSalt, le32(i), blockSize), seg.subarray(0, len)), i * 4096);
  }
  return out.subarray(0, size);
}

const DATASPACE_TRANSFORM = "{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}";
/** A length-prefixed UTF-16 string, padded to 4 bytes (MS-OFFCRYPTO's data spaces). */
function lpString(s: string): Uint8Array {
  const body = new Uint8Array(s.length * 2);
  for (let i = 0; i < s.length; i++) { body[i * 2] = s.charCodeAt(i) & 0xff; body[i * 2 + 1] = s.charCodeAt(i) >> 8; }
  const pad = (4 - (body.length % 4)) % 4;
  return cat(le32(body.length), body, new Uint8Array(pad));
}

/** The data-space streams Word writes beside an encrypted package. */
function dataSpaces(): Map<string, Uint8Array> {
  // Reader, updater and writer versions: 1.0 each (major, minor as 16-bit numbers).
  const ver = new Uint8Array([1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0]);
  const version = cat(lpString("Microsoft.Container.DataSpaces"), ver);
  const entry = cat(le32(1), le32(0), lpString("EncryptedPackage"), lpString("StrongEncryptionDataSpace"));
  const map = cat(le32(8), le32(1), le32(entry.length + 4), entry);
  const info = cat(le32(8), le32(1), lpString("StrongEncryptionTransform"));
  const id = lpString(DATASPACE_TRANSFORM);
  const primary = cat(le32(8 + id.length), le32(1), id, lpString("Microsoft.Container.EncryptionTransform"), ver, le32(0), le32(0), le32(0), le32(4));
  return new Map([
    ["\u0006DataSpaces/Version", version],
    ["\u0006DataSpaces/DataSpaceMap", map],
    ["\u0006DataSpaces/DataSpaceInfo/StrongEncryptionDataSpace", info],
    ["\u0006DataSpaces/TransformInfo/StrongEncryptionTransform/\u0006Primary", primary],
  ]);
}

/** A .docx protected with a password (Word's Agile Encryption: AES-256, SHA-512, 100,000 rounds). */
export async function encryptDocx(docx: Uint8Array, password: string, spins = 100000): Promise<Uint8Array> {
  const keyBytes = 32;
  const kSalt = random(16);
  const pSalt = random(16);
  const secret = random(keyBytes);
  // The package, in 4096-byte segments, each with its own IV; the last padded to the block size.
  const segments: Uint8Array[] = [];
  for (let i = 0, o = 0; o < docx.length; i++, o += 4096) {
    const seg = docx.subarray(o, Math.min(docx.length, o + 4096));
    const padded = new Uint8Array(Math.ceil(seg.length / 16) * 16);
    padded.set(seg);
    segments.push(await aes(true, secret, await ivFor(kSalt, le32(i), 16), padded));
  }
  const size = new Uint8Array(8);
  new DataView(size.buffer).setBigUint64(0, BigInt(docx.length), true);
  const pkg = cat(size, ...segments);
  // Integrity: an HMAC of the encrypted package, with its key and value encrypted.
  const hmacKey = random(64);
  const hk = await subtle().importKey("raw", hmacKey as Uint8Array<ArrayBuffer>, { name: "HMAC", hash: "SHA-512" }, false, ["sign"]);
  const hmac = new Uint8Array(await subtle().sign("HMAC", hk, pkg as Uint8Array<ArrayBuffer>));
  const encHmacKey = await aes(true, secret, await ivFor(kSalt, BLOCK.integrityKey, 16), hmacKey);
  const encHmacValue = await aes(true, secret, await ivFor(kSalt, BLOCK.integrityValue, 16), hmac);
  // The password: a verifier (random input and its hash) and the secret key, each encrypted under a key from it.
  const h = await passwordHash(password, pSalt, spins);
  const verifier = random(16);
  const encInput = await aes(true, await keyFor(h, BLOCK.verifierInput, keyBytes), pSalt, verifier);
  const encHash = await aes(true, await keyFor(h, BLOCK.verifierHash, keyBytes), pSalt, await sha512(verifier));
  const encKey = await aes(true, await keyFor(h, BLOCK.keyValue, keyBytes), pSalt, secret);
  const common = `saltSize="16" blockSize="16" keyBits="256" hashSize="64" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA512"`;
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password" xmlns:c="http://schemas.microsoft.com/office/2006/keyEncryptor/certificate"><keyData ${common} saltValue="${b64(kSalt)}"/><dataIntegrity encryptedHmacKey="${b64(encHmacKey)}" encryptedHmacValue="${b64(encHmacValue)}"/><keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password"><p:encryptedKey spinCount="${spins}" ${common} saltValue="${b64(pSalt)}" encryptedVerifierHashInput="${b64(encInput)}" encryptedVerifierHashValue="${b64(encHash)}" encryptedKeyValue="${b64(encKey)}"/></keyEncryptor></keyEncryptors></encryption>`;
  const info = cat(new Uint8Array([4, 0, 4, 0, 0x40, 0, 0, 0]), new TextEncoder().encode(xml));
  const streams = dataSpaces();
  streams.set("EncryptionInfo", info);
  streams.set("EncryptedPackage", pkg);
  return writeCfb(streams);
}
