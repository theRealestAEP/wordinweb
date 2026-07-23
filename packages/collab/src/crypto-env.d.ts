/**
 * Minimal WebCrypto ambient types for a lib:ES2020 (DOM-free) package.
 * The runtime globals exist in browsers and Node >=16; only the TYPES are
 * missing without the DOM lib, and pulling all of DOM into a deliberately
 * environment-agnostic package for three interfaces would be backwards.
 */
declare type BufferSource = ArrayBufferView | ArrayBuffer;
declare interface CryptoKey {
  readonly type: string;
}
declare interface SubtleCrypto {
  importKey(format: string, keyData: BufferSource, algorithm: unknown, extractable: boolean, keyUsages: string[]): Promise<CryptoKey>;
  deriveKey(algorithm: unknown, baseKey: CryptoKey, derivedKeyType: unknown, extractable: boolean, keyUsages: string[]): Promise<CryptoKey>;
  deriveBits(algorithm: unknown, baseKey: CryptoKey, length: number): Promise<ArrayBuffer>;
  encrypt(algorithm: unknown, key: CryptoKey, data: BufferSource): Promise<ArrayBuffer>;
  decrypt(algorithm: unknown, key: CryptoKey, data: BufferSource): Promise<ArrayBuffer>;
  digest(algorithm: string, data: BufferSource): Promise<ArrayBuffer>;
}
declare const crypto: {
  getRandomValues<T extends ArrayBufferView>(array: T): T;
  subtle: SubtleCrypto;
};
declare function btoa(data: string): string;
declare function atob(data: string): string;
