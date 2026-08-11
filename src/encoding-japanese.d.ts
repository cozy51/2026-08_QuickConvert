declare module 'encoding-japanese' {
  type ConvertOptions = { to: string; from?: string; type?: 'array' | 'string' };
  const Encoding: {
    stringToCode(value: string): number[];
    convert(value: number[] | Uint8Array, options: ConvertOptions): number[];
  };
  export default Encoding;
}
