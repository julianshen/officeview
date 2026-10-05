// Modified by Officeview: bounded allocations (see NOTICE.md).
const OFFICEVIEW_MAX_BYTES = 32 * 1024 * 1024;
let officeviewByteLimit = OFFICEVIEW_MAX_BYTES;
export function withDecoderByteLimit(limit, decode) {
  if (!Number.isSafeInteger(limit) || limit < 0 || limit > OFFICEVIEW_MAX_BYTES) throw new Error("Invalid font decoder budget");
  const previous = officeviewByteLimit;
  officeviewByteLimit = Math.min(previous, limit);
  try { return decode(); } finally { officeviewByteLimit = previous; }
}
function boundedLength(n, width = 1) {
  if (!Number.isSafeInteger(n) || n < 0 || n * width > officeviewByteLimit) {
    throw new Error("Font decoder allocation budget exceeded");
  }
  return n;
}
// src/errors.ts
var EOT_WARN = 1e3;
var EotErrorCode = /* @__PURE__ */ ((EotErrorCode2) => {
  EotErrorCode2["InsufficientBytes"] = "INSUFFICIENT_BYTES";
  EotErrorCode2["HeaderTooBig"] = "HEADER_TOO_BIG";
  EotErrorCode2["BogusStringSize"] = "BOGUS_STRING_SIZE";
  EotErrorCode2["CorruptFile"] = "CORRUPT_FILE";
  EotErrorCode2["LogicError"] = "LOGIC_ERROR";
  EotErrorCode2["NoMaxpTable"] = "NO_MAXP_TABLE";
  EotErrorCode2["NoHeadTable"] = "NO_HEAD_TABLE";
  EotErrorCode2["NoHmtxTable"] = "NO_HMTX_TABLE";
  EotErrorCode2["CorruptHopcodeData"] = "CORRUPT_HOPCODE_DATA";
  EotErrorCode2["MalformedHeadTable"] = "MALFORMED_HEAD_TABLE";
  EotErrorCode2["OffByteBoundary"] = "OFF_BYTE_BOUNDARY";
  EotErrorCode2["OutOfReservedSpace"] = "OUT_OF_RESERVED_SPACE";
  EotErrorCode2["SeekPastEos"] = "SEEK_PAST_EOS";
  EotErrorCode2["MtxError"] = "MTX_ERROR";
  EotErrorCode2["WarnBadVersion"] = "WARN_BAD_VERSION";
  return EotErrorCode2;
})(EotErrorCode || {});
var WARNING_CODES = /* @__PURE__ */ new Set(["WARN_BAD_VERSION" /* WarnBadVersion */]);
var EotError = class _EotError extends Error {
  code;
  constructor(code, message) {
    super(message);
    this.name = "EotError";
    this.code = code;
    Object.setPrototypeOf(this, _EotError.prototype);
  }
  /** True when this represents a recoverable warning rather than a fatal error. */
  get isWarning() {
    return WARNING_CODES.has(this.code);
  }
};

// src/stream.ts
var Stream = class _Stream {
  buf;
  size;
  // how much data has been written or is valid
  reserved;
  // allocated capacity
  pos;
  // current byte position
  bitPos;
  // current bit position within the byte at `pos`
  constructor(buf, size) {
    if (buf) {
      this.buf = buf;
      this.size = size;
      this.reserved = buf.length;
    } else {
      this.buf = new Uint8Array(boundedLength(0));
      this.size = 0;
      this.reserved = 0;
    }
    this.pos = 0;
    this.bitPos = 0;
  }
  static fromExisting(buf, size, reserved) {
    const s = new _Stream(null, 0);
    s.buf = buf;
    s.size = size;
    s.reserved = reserved;
    return s;
  }
  reserve(n) {
    boundedLength(n);
    if (this.reserved >= n) {
      return;
    }
    const newBuf = new Uint8Array(boundedLength(n));
    newBuf.set(this.buf.subarray(0, this.size));
    this.buf = newBuf;
    this.reserved = n;
  }
  /**
   * Reject byte-level access while the stream sits mid-byte (`bitPos != 0`).
   *
   * libeot returns `EOT_OFF_BYTE_BOUNDARY` for any byte read/write/seek issued
   * before a partial byte has been consumed. The bit-level reader
   * ({@link readNBits}) is the sole legitimate mid-byte accessor and bypasses
   * this guard by touching `buf`/`pos` directly. On valid input the only
   * `readNBits` caller consumes whole bytes per point, so the stream is always
   * byte-aligned when a byte accessor runs and this guard never fires.
   */
  ensureByteAligned() {
    if (this.bitPos !== 0) {
      throw new EotError(
        "OFF_BYTE_BOUNDARY" /* OffByteBoundary */,
        `Stream: byte-level access at a non-byte boundary (bitPos=${this.bitPos}, pos=${this.pos})`
      );
    }
  }
  ensureWrite(n) {
    this.ensureByteAligned();
    const needed = this.pos + n;
    if (needed > this.reserved) {
      this.reserve(Math.max(needed, this.reserved * 2 || 256));
    }
    if (needed > this.size) {
      this.size = needed;
    }
  }
  ensureRead(n) {
    this.ensureByteAligned();
    if (this.pos + n > this.size) {
      throw new EotError(
        "INSUFFICIENT_BYTES" /* InsufficientBytes */,
        `Stream: not enough data (need ${n} bytes at pos ${this.pos}, size ${this.size})`
      );
    }
  }
  // --- Seek ---
  // A seek requires the stream to be byte-aligned and never clears `bitPos`
  // itself (mirroring libeot, which refuses to seek mid-byte rather than
  // silently re-aligning). The alignment guard leaves `bitPos` at 0.
  seekAbsolute(pos) {
    this.ensureByteAligned();
    if (!Number.isInteger(pos) || pos < 0) {
      throw new EotError("CORRUPT_FILE" /* CorruptFile */, `Stream: invalid seek position (${pos})`);
    }
    if (pos > this.size) {
      throw new EotError("SEEK_PAST_EOS" /* SeekPastEos */, `Stream: seek past end (${pos} > ${this.size})`);
    }
    this.pos = pos;
  }
  seekRelative(offset) {
    this.ensureByteAligned();
    const newPos = this.pos + offset;
    if (!Number.isInteger(offset) || !Number.isInteger(newPos)) {
      throw new EotError("CORRUPT_FILE" /* CorruptFile */, `Stream: invalid relative seek (${offset})`);
    }
    if (newPos < 0) {
      throw new EotError("CORRUPT_FILE" /* CorruptFile */, "Stream: negative seek");
    }
    if (newPos > this.size) {
      throw new EotError("SEEK_PAST_EOS" /* SeekPastEos */, "Stream: seek past end");
    }
    this.pos = newPos;
  }
  // Seek into already-reserved-but-unwritten space, extending `size` up to the
  // seek target. libeot returns `EOT_SEEK_PAST_EOS` when the target exceeds the
  // reserved capacity; we mirror that rather than growing the buffer, so an
  // over-reach surfaces as a failure instead of a silent realloc.
  seekAbsoluteThroughReserve(pos) {
    this.ensureByteAligned();
    if (pos > this.reserved) {
      throw new EotError(
        "SEEK_PAST_EOS" /* SeekPastEos */,
        `Stream: seek to ${pos} past reserved end (${this.reserved})`
      );
    }
    if (pos > this.size) {
      this.size = pos;
    }
    this.pos = pos;
  }
  seekRelativeThroughReserve(offset) {
    this.seekAbsoluteThroughReserve(this.pos + offset);
  }
  // --- Read (Big-Endian) ---
  readU8() {
    this.ensureRead(1);
    return this.buf[this.pos++];
  }
  peekU8() {
    this.ensureRead(1);
    return this.buf[this.pos];
  }
  readU16() {
    this.ensureRead(2);
    const v = this.buf[this.pos] << 8 | this.buf[this.pos + 1];
    this.pos += 2;
    return v;
  }
  readU24() {
    this.ensureRead(3);
    const v = this.buf[this.pos] << 16 | this.buf[this.pos + 1] << 8 | this.buf[this.pos + 2];
    this.pos += 3;
    return v;
  }
  readU32() {
    this.ensureRead(4);
    const v = (this.buf[this.pos] << 24 | this.buf[this.pos + 1] << 16 | this.buf[this.pos + 2] << 8 | this.buf[this.pos + 3]) >>> 0;
    this.pos += 4;
    return v;
  }
  readS16() {
    const v = this.readU16();
    return v >= 32768 ? v - 65536 : v;
  }
  readS8() {
    const v = this.readU8();
    return v >= 128 ? v - 256 : v;
  }
  readChar() {
    return String.fromCharCode(this.readU8());
  }
  // --- Write (Big-Endian) ---
  writeU8(v) {
    this.ensureWrite(1);
    this.buf[this.pos++] = v & 255;
  }
  writeU16(v) {
    this.ensureWrite(2);
    this.buf[this.pos++] = v >> 8 & 255;
    this.buf[this.pos++] = v & 255;
  }
  writeU24(v) {
    if (v < 0 || v > 16777215) {
      throw new Error(`Stream: writeU24 value out of range: ${v}`);
    }
    this.ensureWrite(3);
    this.buf[this.pos++] = v >> 16 & 255;
    this.buf[this.pos++] = v >> 8 & 255;
    this.buf[this.pos++] = v & 255;
  }
  writeU32(v) {
    this.ensureWrite(4);
    this.buf[this.pos++] = v >>> 24 & 255;
    this.buf[this.pos++] = v >> 16 & 255;
    this.buf[this.pos++] = v >> 8 & 255;
    this.buf[this.pos++] = v & 255;
  }
  writeS16(v) {
    this.writeU16(v < 0 ? v + 65536 : v);
  }
  writeS8(v) {
    this.writeU8(v < 0 ? v + 256 : v);
  }
  // --- Bit-level reading (for triplet coordinate decoding) ---
  readNBits(n) {
    if (n === 0) {
      return 0;
    }
    if (n > 32) {
      throw new Error(`Stream: readNBits width out of range: ${n} (max 32)`);
    }
    let value = 0;
    let bitsRemaining = n;
    while (bitsRemaining > 0) {
      if (this.pos >= this.size && this.bitPos === 0) {
        throw new EotError("INSUFFICIENT_BYTES" /* InsufficientBytes */, "Stream: not enough data for bit read");
      }
      const bitsAvailableInByte = 8 - this.bitPos;
      const bitsToRead = Math.min(bitsRemaining, bitsAvailableInByte);
      const shift = bitsAvailableInByte - bitsToRead;
      const mask = (1 << bitsToRead) - 1 << shift;
      value = value << bitsToRead | (this.buf[this.pos] & mask) >> shift;
      this.bitPos += bitsToRead;
      if (this.bitPos >= 8) {
        this.bitPos = 0;
        this.pos++;
      }
      bitsRemaining -= bitsToRead;
    }
    return value >>> 0;
  }
  // --- Copy ---
  /**
   * Copy `length` bytes from this stream to `dest`.
   *
   * Both streams must be byte-aligned. The destination must already have the
   * capacity reserved: libeot returns `EOT_OUT_OF_RESERVED_SPACE` when a copy
   * would overrun the reserved buffer, so we throw rather than auto-growing —
   * an under-reservation is a bug we want surfaced, not silently patched.
   */
  copyTo(dest, length) {
    this.ensureByteAligned();
    dest.ensureByteAligned();
    if (this.pos + length > this.size) {
      throw new EotError("INSUFFICIENT_BYTES" /* InsufficientBytes */, "Stream: not enough data for copy");
    }
    const needed = dest.pos + length;
    if (needed > dest.reserved) {
      throw new EotError(
        "OUT_OF_RESERVED_SPACE" /* OutOfReservedSpace */,
        `Stream: copy of ${length} bytes exceeds reserved capacity (need ${needed}, reserved ${dest.reserved})`
      );
    }
    dest.buf.set(this.buf.subarray(this.pos, this.pos + length), dest.pos);
    this.pos += length;
    dest.pos += length;
    if (dest.pos > dest.size) {
      dest.size = dest.pos;
    }
  }
  /** Read rest of data as 4-byte-aligned U32 values. Returns 0 on incomplete read. */
  readRestAsU32() {
    if (this.pos + 4 > this.size) {
      if (this.pos >= this.size) {
        return null;
      }
      let val = 0;
      const remaining = this.size - this.pos;
      for (let i = 0; i < 4; i++) {
        val <<= 8;
        if (i < remaining) {
          val |= this.buf[this.pos + i];
        }
      }
      this.pos = this.size;
      return val >>> 0;
    }
    return this.readU32();
  }
  /**
   * Compute the SFNT-style checksum of bytes in `[beginPos, endPos)` as a sum
   * of big-endian U32 words, zero-padding a final partial word. Bounds strictly
   * on `endPos` (not the stream's `size`), so an unaligned range never folds in
   * bytes past `endPos`. Leaves the stream position unchanged.
   */
  checksumU32(beginPos, endPos) {
    if (beginPos > endPos) {
      throw new Error(`Stream: checksumU32 beginPos ${beginPos} > endPos ${endPos}`);
    }
    if (endPos > this.size) {
      throw new Error(`Stream: checksumU32 endPos ${endPos} exceeds size ${this.size}`);
    }
    let sum = 0;
    for (let p = beginPos; p < endPos; p += 4) {
      let word = 0;
      for (let i = 0; i < 4; i++) {
        word = word << 8 | (p + i < endPos ? this.buf[p + i] : 0);
      }
      sum = sum + (word >>> 0) >>> 0;
    }
    return sum;
  }
  /** Get a copy of the written data. */
  toUint8Array() {
    return this.buf.slice(0, this.size);
  }
};

// src/triplet-encodings.ts
var TRIPLET_ENCODINGS = [
  // Indices 0-9: xBits=0 or yBits=0 (one axis)
  { byteCount: 2, xBits: 0, yBits: 8, deltaX: 0, deltaY: 0, xSign: 0, ySign: -1 },
  { byteCount: 2, xBits: 0, yBits: 8, deltaX: 0, deltaY: 0, xSign: 0, ySign: 1 },
  { byteCount: 2, xBits: 0, yBits: 8, deltaX: 0, deltaY: 256, xSign: 0, ySign: -1 },
  { byteCount: 2, xBits: 0, yBits: 8, deltaX: 0, deltaY: 256, xSign: 0, ySign: 1 },
  { byteCount: 2, xBits: 0, yBits: 8, deltaX: 0, deltaY: 512, xSign: 0, ySign: -1 },
  { byteCount: 2, xBits: 0, yBits: 8, deltaX: 0, deltaY: 512, xSign: 0, ySign: 1 },
  { byteCount: 2, xBits: 0, yBits: 8, deltaX: 0, deltaY: 768, xSign: 0, ySign: -1 },
  { byteCount: 2, xBits: 0, yBits: 8, deltaX: 0, deltaY: 768, xSign: 0, ySign: 1 },
  { byteCount: 2, xBits: 0, yBits: 8, deltaX: 0, deltaY: 1024, xSign: 0, ySign: -1 },
  { byteCount: 2, xBits: 0, yBits: 8, deltaX: 0, deltaY: 1024, xSign: 0, ySign: 1 },
  // Indices 10-19: yBits=0 (X axis only)
  { byteCount: 2, xBits: 8, yBits: 0, deltaX: 0, deltaY: 0, xSign: -1, ySign: 0 },
  { byteCount: 2, xBits: 8, yBits: 0, deltaX: 0, deltaY: 0, xSign: 1, ySign: 0 },
  { byteCount: 2, xBits: 8, yBits: 0, deltaX: 256, deltaY: 0, xSign: -1, ySign: 0 },
  { byteCount: 2, xBits: 8, yBits: 0, deltaX: 256, deltaY: 0, xSign: 1, ySign: 0 },
  { byteCount: 2, xBits: 8, yBits: 0, deltaX: 512, deltaY: 0, xSign: -1, ySign: 0 },
  { byteCount: 2, xBits: 8, yBits: 0, deltaX: 512, deltaY: 0, xSign: 1, ySign: 0 },
  { byteCount: 2, xBits: 8, yBits: 0, deltaX: 768, deltaY: 0, xSign: -1, ySign: 0 },
  { byteCount: 2, xBits: 8, yBits: 0, deltaX: 768, deltaY: 0, xSign: 1, ySign: 0 },
  { byteCount: 2, xBits: 8, yBits: 0, deltaX: 1024, deltaY: 0, xSign: -1, ySign: 0 },
  { byteCount: 2, xBits: 8, yBits: 0, deltaX: 1024, deltaY: 0, xSign: 1, ySign: 0 },
  // Indices 20-83: 4-bit X + 4-bit Y (2 bytes total)
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 1, deltaY: 1, xSign: -1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 1, deltaY: 1, xSign: 1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 1, deltaY: 1, xSign: -1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 1, deltaY: 1, xSign: 1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 1, deltaY: 17, xSign: -1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 1, deltaY: 17, xSign: 1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 1, deltaY: 17, xSign: -1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 1, deltaY: 17, xSign: 1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 1, deltaY: 33, xSign: -1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 1, deltaY: 33, xSign: 1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 1, deltaY: 33, xSign: -1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 1, deltaY: 33, xSign: 1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 1, deltaY: 49, xSign: -1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 1, deltaY: 49, xSign: 1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 1, deltaY: 49, xSign: -1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 1, deltaY: 49, xSign: 1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 17, deltaY: 1, xSign: -1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 17, deltaY: 1, xSign: 1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 17, deltaY: 1, xSign: -1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 17, deltaY: 1, xSign: 1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 17, deltaY: 17, xSign: -1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 17, deltaY: 17, xSign: 1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 17, deltaY: 17, xSign: -1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 17, deltaY: 17, xSign: 1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 17, deltaY: 33, xSign: -1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 17, deltaY: 33, xSign: 1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 17, deltaY: 33, xSign: -1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 17, deltaY: 33, xSign: 1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 17, deltaY: 49, xSign: -1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 17, deltaY: 49, xSign: 1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 17, deltaY: 49, xSign: -1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 17, deltaY: 49, xSign: 1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 33, deltaY: 1, xSign: -1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 33, deltaY: 1, xSign: 1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 33, deltaY: 1, xSign: -1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 33, deltaY: 1, xSign: 1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 33, deltaY: 17, xSign: -1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 33, deltaY: 17, xSign: 1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 33, deltaY: 17, xSign: -1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 33, deltaY: 17, xSign: 1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 33, deltaY: 33, xSign: -1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 33, deltaY: 33, xSign: 1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 33, deltaY: 33, xSign: -1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 33, deltaY: 33, xSign: 1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 33, deltaY: 49, xSign: -1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 33, deltaY: 49, xSign: 1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 33, deltaY: 49, xSign: -1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 33, deltaY: 49, xSign: 1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 49, deltaY: 1, xSign: -1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 49, deltaY: 1, xSign: 1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 49, deltaY: 1, xSign: -1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 49, deltaY: 1, xSign: 1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 49, deltaY: 17, xSign: -1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 49, deltaY: 17, xSign: 1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 49, deltaY: 17, xSign: -1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 49, deltaY: 17, xSign: 1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 49, deltaY: 33, xSign: -1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 49, deltaY: 33, xSign: 1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 49, deltaY: 33, xSign: -1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 49, deltaY: 33, xSign: 1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 49, deltaY: 49, xSign: -1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 49, deltaY: 49, xSign: 1, ySign: -1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 49, deltaY: 49, xSign: -1, ySign: 1 },
  { byteCount: 2, xBits: 4, yBits: 4, deltaX: 49, deltaY: 49, xSign: 1, ySign: 1 },
  // Indices 84-119: 8-bit X + 8-bit Y (3 bytes total)
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 1, deltaY: 1, xSign: -1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 1, deltaY: 1, xSign: 1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 1, deltaY: 1, xSign: -1, ySign: 1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 1, deltaY: 1, xSign: 1, ySign: 1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 1, deltaY: 257, xSign: -1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 1, deltaY: 257, xSign: 1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 1, deltaY: 257, xSign: -1, ySign: 1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 1, deltaY: 257, xSign: 1, ySign: 1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 1, deltaY: 513, xSign: -1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 1, deltaY: 513, xSign: 1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 1, deltaY: 513, xSign: -1, ySign: 1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 1, deltaY: 513, xSign: 1, ySign: 1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 257, deltaY: 1, xSign: -1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 257, deltaY: 1, xSign: 1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 257, deltaY: 1, xSign: -1, ySign: 1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 257, deltaY: 1, xSign: 1, ySign: 1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 257, deltaY: 257, xSign: -1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 257, deltaY: 257, xSign: 1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 257, deltaY: 257, xSign: -1, ySign: 1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 257, deltaY: 257, xSign: 1, ySign: 1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 257, deltaY: 513, xSign: -1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 257, deltaY: 513, xSign: 1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 257, deltaY: 513, xSign: -1, ySign: 1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 257, deltaY: 513, xSign: 1, ySign: 1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 513, deltaY: 1, xSign: -1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 513, deltaY: 1, xSign: 1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 513, deltaY: 1, xSign: -1, ySign: 1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 513, deltaY: 1, xSign: 1, ySign: 1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 513, deltaY: 257, xSign: -1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 513, deltaY: 257, xSign: 1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 513, deltaY: 257, xSign: -1, ySign: 1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 513, deltaY: 257, xSign: 1, ySign: 1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 513, deltaY: 513, xSign: -1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 513, deltaY: 513, xSign: 1, ySign: -1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 513, deltaY: 513, xSign: -1, ySign: 1 },
  { byteCount: 3, xBits: 8, yBits: 8, deltaX: 513, deltaY: 513, xSign: 1, ySign: 1 },
  // Indices 120-123: 12-bit X + 12-bit Y (4 bytes total)
  { byteCount: 4, xBits: 12, yBits: 12, deltaX: 0, deltaY: 0, xSign: -1, ySign: -1 },
  { byteCount: 4, xBits: 12, yBits: 12, deltaX: 0, deltaY: 0, xSign: 1, ySign: -1 },
  { byteCount: 4, xBits: 12, yBits: 12, deltaX: 0, deltaY: 0, xSign: -1, ySign: 1 },
  { byteCount: 4, xBits: 12, yBits: 12, deltaX: 0, deltaY: 0, xSign: 1, ySign: 1 },
  // Indices 124-127: 16-bit X + 16-bit Y (5 bytes total)
  { byteCount: 5, xBits: 16, yBits: 16, deltaX: 0, deltaY: 0, xSign: -1, ySign: -1 },
  { byteCount: 5, xBits: 16, yBits: 16, deltaX: 0, deltaY: 0, xSign: 1, ySign: -1 },
  { byteCount: 5, xBits: 16, yBits: 16, deltaX: 0, deltaY: 0, xSign: -1, ySign: 1 },
  { byteCount: 5, xBits: 16, yBits: 16, deltaX: 0, deltaY: 0, xSign: 1, ySign: 1 }
];

// src/magnitude.ts
function readMagnitude(stream) {
  const readBit = () => {
    if (stream.pos >= stream.size) {
      throw new EotError("INSUFFICIENT_BYTES" /* InsufficientBytes */, "truncated magnitude value");
    }
    const bit = stream.buf[stream.pos] >> stream.bitPos & 1;
    stream.bitPos++;
    if (stream.bitPos === 8) {
      stream.bitPos = 0;
      stream.pos++;
    }
    return bit;
  };
  if (readBit() === 0) return 0;
  let magnitude = 1;
  while (readBit() === 1) magnitude++;
  return readBit() === 1 ? -magnitude : magnitude;
}

// src/hdmx.ts
var MAX_OUTPUT_BYTES = 32 * 1024 * 1024;
function fail(message) {
  throw new EotError("CORRUPT_FILE" /* CorruptFile */, `invalid hdmx table: ${message}`);
}
function u16(bytes, offset) {
  return bytes[offset] << 8 | bytes[offset + 1];
}
function u32(bytes, offset) {
  return (bytes[offset] << 24 | bytes[offset + 1] << 16 | bytes[offset + 2] << 8 | bytes[offset + 3]) >>> 0;
}
function putU16(bytes, offset, value) {
  bytes[offset] = value >>> 8;
  bytes[offset + 1] = value & 255;
}
function putU32(bytes, offset, value) {
  bytes[offset] = value >>> 24;
  bytes[offset + 1] = value >>> 16;
  bytes[offset + 2] = value >>> 8;
  bytes[offset + 3] = value;
}
function decodeHdmx(data, metadata) {
  if (data.length < 8) {
    throw new EotError("INSUFFICIENT_BYTES" /* InsufficientBytes */, "truncated hdmx header");
  }
  const encodedVersion = u16(data, 0);
  const numRecords = u16(data, 2);
  const recordSize = u32(data, 4);
  if (encodedVersion === 65535) {
    if (recordSize < 2 || recordSize % 4 !== 0) fail("invalid raw record size");
    const rawLength = 8 + numRecords * recordSize;
    if (!Number.isSafeInteger(rawLength) || rawLength > data.length) {
      throw new EotError("INSUFFICIENT_BYTES" /* InsufficientBytes */, "truncated raw hdmx records");
    }
    const raw = data.slice();
    putU16(raw, 0, 0);
    return raw;
  }
  if (encodedVersion !== 0) fail(`unsupported version ${encodedVersion}`);
  if (!metadata) fail("missing head/hhea/hmtx/maxp metadata");
  const { numGlyphs, unitsPerEm, numberOfHMetrics, hmtx } = metadata;
  if (!Number.isInteger(numGlyphs) || numGlyphs < 0 || numGlyphs > 65535) fail("invalid numGlyphs");
  if (!Number.isInteger(unitsPerEm) || unitsPerEm <= 0 || unitsPerEm > 65535) fail("invalid unitsPerEm");
  if (!Number.isInteger(numberOfHMetrics) || numberOfHMetrics < 1 || numberOfHMetrics > numGlyphs) {
    fail("invalid numberOfHMetrics");
  }
  const expectedRecordSize = numGlyphs + 2 + 3 & -4;
  if (recordSize !== expectedRecordSize) fail("recordSize does not match glyph count");
  const outputLength = 8 + numRecords * recordSize;
  if (!Number.isSafeInteger(outputLength) || outputLength > MAX_OUTPUT_BYTES) {
    fail("decoded table is too large");
  }
  const headerLength = 8 + numRecords * 2;
  if (headerLength > data.length) {
    throw new EotError("INSUFFICIENT_BYTES" /* InsufficientBytes */, "truncated hdmx record headers");
  }
  const minBitsBytes = Math.ceil(numRecords * numGlyphs / 8);
  if (headerLength + minBitsBytes > data.length) {
    throw new EotError("INSUFFICIENT_BYTES" /* InsufficientBytes */, "truncated hdmx magnitude data");
  }
  const hmtxLength = numberOfHMetrics * 4 + (numGlyphs - numberOfHMetrics) * 2;
  if (hmtx.length < hmtxLength) fail("hmtx table is too short");
  const advances = new Uint16Array(boundedLength(numGlyphs, 2));
  let lastAdvance = 0;
  for (let glyph = 0; glyph < numGlyphs; glyph++) {
    if (glyph < numberOfHMetrics) {
      lastAdvance = u16(hmtx, glyph * 4);
    }
    advances[glyph] = lastAdvance;
  }
  const out = new Uint8Array(boundedLength(outputLength));
  putU16(out, 0, encodedVersion);
  putU16(out, 2, numRecords);
  putU32(out, 4, recordSize);
  const bits = new Stream(data.subarray(headerLength), data.length - headerLength);
  for (let record = 0; record < numRecords; record++) {
    const sourceRecord = 8 + record * 2;
    const targetRecord = 8 + record * recordSize;
    const ppem = data[sourceRecord];
    out[targetRecord] = ppem;
    out[targetRecord + 1] = data[sourceRecord + 1];
    for (let glyph = 0; glyph < numGlyphs; glyph++) {
      const aw = advances[glyph];
      const rounded64 = Math.floor((64 * ppem * aw + unitsPerEm / 2) / unitsPerEm);
      const predicted = Math.floor((rounded64 + 32) / 64);
      const width = predicted + readMagnitude(bits);
      if (width < 0 || width > 255) fail(`decoded width out of range (${width})`);
      out[targetRecord + 2 + glyph] = width;
    }
  }
  return out;
}

// src/vdmx.ts
function malformed(message) {
  throw new EotError("CORRUPT_FILE" /* CorruptFile */, `malformed VDMX table: ${message}`);
}
function ensureRange(start, length, limit, what) {
  if (!Number.isInteger(start) || !Number.isInteger(length) || start < 0 || length < 0 || start + length > limit) {
    malformed(`${what} is outside the ${limit}-byte table`);
  }
}
function toInt16(value, what) {
  if (!Number.isInteger(value) || value < -32768 || value > 32767) {
    malformed(`${what} is outside the signed 16-bit range`);
  }
  return value;
}
function validateUncompressed(data) {
  if (data.length < 6) malformed("table is shorter than its header");
  const input = new Stream(data, data.length);
  const version = input.readU16();
  if (version !== 0 && version !== 1) malformed(`unsupported OpenType version ${version}`);
  const numRecs = input.readU16();
  const numRatios = input.readU16();
  if (numRecs === 0 || numRatios === 0) malformed("table must contain at least one group and ratio");
  const headerEnd = 6 + numRatios * 6;
  ensureRange(6, numRatios * 4, data.length, "ratio records");
  input.seekAbsolute(6 + numRatios * 4);
  const offsets = [];
  for (let i = 0; i < numRatios; i++) offsets.push(input.readU16());
  const groupStarts = [...new Set(offsets)].sort((a, b) => a - b);
  if (groupStarts.length !== numRecs) {
    malformed(`ratio offsets identify ${groupStarts.length} groups, header declares ${numRecs}`);
  }
  for (let i = 0; i < groupStarts.length; i++) {
    const start = groupStarts[i];
    if (start < headerEnd) malformed("group offset overlaps the header");
    ensureRange(start, 4, data.length, `group ${i} header`);
    input.seekAbsolute(start);
    const recs = input.readU16();
    input.readU8();
    input.readU8();
    if (recs === 0) malformed(`group ${i} has no records`);
    const size = 4 + recs * 6;
    const limit = groupStarts[i + 1] ?? data.length;
    if (start + size > limit) malformed(`group ${i} extends past its table boundary`);
    let previousPpem = -1;
    for (let j = 0; j < recs; j++) {
      const ppem = input.readU16();
      input.readS16();
      input.readS16();
      if (ppem <= previousPpem) malformed(`group ${i} has unsorted ppem heights`);
      previousPpem = ppem;
    }
  }
}
function decodeVdmx(data) {
  if (data.length < 2) {
    throw new EotError("INSUFFICIENT_BYTES" /* InsufficientBytes */, "VDMX table is shorter than its version field");
  }
  const input = new Stream(data, data.length);
  const version = input.readU16();
  if (version > 1) {
    const originalVersion = 65535 - version;
    if (originalVersion !== 0 && originalVersion !== 1) {
      malformed(`unsupported version marker ${version}`);
    }
    const output2 = data.slice();
    output2[0] = originalVersion >>> 8;
    output2[1] = originalVersion & 255;
    validateUncompressed(output2);
    return output2;
  }
  ensureRange(0, 6, data.length, "header");
  const numRecs = input.readU16();
  const numRatios = input.readU16();
  if (numRecs === 0 || numRatios === 0) {
    malformed("compressed table must contain at least one group and ratio");
  }
  const headerEnd = 6 + numRatios * 6;
  ensureRange(6, numRatios * 4, data.length, "ratio records");
  input.seekAbsolute(6 + numRatios * 4);
  const offsets = [];
  for (let i = 0; i < numRatios; i++) offsets.push(input.readU16());
  const groupStarts = [...new Set(offsets)].sort((a, b) => a - b);
  if (groupStarts.length !== numRecs) {
    malformed(`ratio offsets identify ${groupStarts.length} groups, header declares ${numRecs}`);
  }
  if (groupStarts[0] !== offsets[0]) {
    malformed("the first ratio offset must point to the first compressed group");
  }
  for (const offset of groupStarts) {
    if (offset < headerEnd) malformed("group offset overlaps the header");
  }
  input.seekAbsolute(offsets[0]);
  let output = data.slice();
  for (let groupIndex = 0; groupIndex < numRecs; groupIndex++) {
    const compressedStart = input.pos;
    const recs = input.readU16();
    const yMaxMultiplier = input.readS16();
    const yMinMultiplier = input.readS16();
    if (recs === 0) malformed(`group ${groupIndex} has no records`);
    const expandedSize = 4 + recs * 6;
    const groupStart = groupStarts[groupIndex];
    const groupLimit = groupStarts[groupIndex + 1] ?? data.length;
    if (groupIndex + 1 < numRecs && groupStart + expandedSize > groupLimit) {
      malformed(`expanded group ${groupIndex} overlaps the next group`);
    }
    const records = new Stream(null, 0);
    records.reserve(recs * 6);
    let predictedPpem = 8;
    let previousPpem = -1;
    let startSize = 0;
    let endSize = 0;
    for (let i = 0; i < recs; i++) {
      const ppemError = readMagnitude(input);
      const yMaxError = readMagnitude(input);
      const yMinError = readMagnitude(input);
      const ppem = ppemError + predictedPpem;
      if (!Number.isInteger(ppem) || ppem < 0 || ppem > 65535 || ppem <= previousPpem) {
        malformed(`group ${groupIndex} has an invalid or unsorted ppem`);
      }
      const predictedYMax = Math.trunc((ppem * yMaxMultiplier + 1024) / 2048);
      const predictedYMin = -Math.trunc((ppem * yMinMultiplier + 1024) / 2048);
      const yMax = toInt16(predictedYMax + yMaxError, `group ${groupIndex} yMax`);
      const yMin = toInt16(predictedYMin + yMinError, `group ${groupIndex} yMin`);
      if (i === 0) startSize = ppem;
      if (i === recs - 1) endSize = ppem;
      records.writeU16(ppem);
      records.writeS16(yMax);
      records.writeS16(yMin);
      previousPpem = ppem;
      predictedPpem = ppem + 1;
    }
    if (input.bitPos !== 0) {
      input.pos++;
      input.bitPos = 0;
    }
    ensureRange(compressedStart, input.pos - compressedStart, data.length, `compressed group ${groupIndex}`);
    const outputGroup = new Stream(null, 0);
    outputGroup.reserve(expandedSize);
    outputGroup.writeU16(recs);
    outputGroup.writeU8(startSize & 255);
    outputGroup.writeU8(endSize & 255);
    for (const byte of records.toUint8Array()) outputGroup.writeU8(byte);
    const groupEnd = groupStart + expandedSize;
    if (groupEnd > output.length) {
      const expanded = new Uint8Array(boundedLength(groupEnd));
      expanded.set(output);
      output = expanded;
    }
    output.set(outputGroup.toUint8Array(), groupStart);
  }
  validateUncompressed(output);
  return output;
}

// src/ctf-parser.ts
var FLG_ON_CURVE = 1;
var FLG_X_SHORT = 2;
var FLG_Y_SHORT = 4;
var FLG_X_SAME = 16;
var FLG_Y_SAME = 32;
var NPUSHB = 64;
var NPUSHW = 65;
var PUSHB = 176;
var PUSHW = 184;
var ARG_1_AND_2_ARE_WORDS = 1;
var HAVE_SCALE = 8;
var MORE_COMPONENTS = 32;
var HAVE_XY_SCALE = 64;
var HAVE_2_BY_2 = 128;
var HAVE_INSTRUCTIONS = 256;
var INT16_MIN = -32768;
var INT16_MAX = 32767;
function toInt162(v) {
  v &= 65535;
  return v >= 32768 ? v - 65536 : v;
}
function read255UShort(s) {
  const code = s.readU8();
  if (code === 253) {
    return s.readU16();
  }
  if (code === 255) {
    return 253 + s.readU8();
  }
  if (code === 254) {
    return 506 + s.readU8();
  }
  return code;
}
function read255Short(s) {
  let sign = 1;
  let code = s.readU8();
  if (code === 253) {
    return s.readS16();
  }
  if (code === 250) {
    sign = -1;
    code = s.readU8();
  }
  let value;
  if (code === 255) {
    value = 250 + s.readU8();
  } else if (code === 254) {
    value = 500 + s.readU8();
  } else {
    value = code;
  }
  return value * sign;
}
function unpackCVT(table, sIn) {
  sIn.seekAbsolute(table.offset);
  const numEntries = sIn.readU16();
  // Include SFNT padding in the remaining aggregate allowance before reserve/copy.
  boundedLength(Math.ceil(numEntries * 2 / 4) * 4);
  const out = new Stream(null, 0);
  out.reserve(numEntries * 2);
  let lastValue = 0;
  for (let i = 0; i < numEntries; i++) {
    const code = sIn.readU8();
    let val;
    if (code >= 248) {
      val = 238 * (code - 247) + sIn.readU8();
    } else if (code >= 239) {
      val = -(238 * (code - 239) + sIn.readU8());
    } else if (code === 238) {
      val = sIn.readS16();
    } else {
      val = code;
    }
    lastValue = toInt162(lastValue + val);
    out.writeS16(lastValue);
  }
  table.buf = out.toUint8Array();
  table.bufSize = table.buf.length;
}
function decodePushInstructions(sIn, sOut, pushCount) {
  if (pushCount === 0) {
    return;
  }
  const data = [];
  let remaining = pushCount;
  let isShort = false;
  const runValues = [];
  function flush() {
    if (runValues.length === 0) {
      return;
    }
    const count = runValues.length;
    if (isShort) {
      if (count < 8) {
        sOut.writeU8(PUSHW + (count - 1));
      } else {
        sOut.writeU8(NPUSHW);
        sOut.writeU8(count);
      }
      for (const v of runValues) {
        sOut.writeS16(v);
      }
    } else {
      if (count < 8) {
        sOut.writeU8(PUSHB + (count - 1));
      } else {
        sOut.writeU8(NPUSHB);
        sOut.writeU8(count);
      }
      for (const v of runValues) {
        sOut.writeU8(v & 255);
      }
    }
    runValues.length = 0;
  }
  function put(v) {
    data.push(v);
    const needsShort = v < 0 || v > 255;
    if (runValues.length > 0 && needsShort !== isShort) {
      flush();
    }
    if (runValues.length === 0) {
      isShort = needsShort;
    }
    runValues.push(v);
    if (runValues.length >= 255) {
      flush();
    }
  }
  while (remaining > 0) {
    const code = sIn.peekU8();
    if (code === 251) {
      if (remaining < 3 || data.length < 2) {
        throw new EotError(
          "CORRUPT_HOPCODE_DATA" /* CorruptHopcodeData */,
          `corrupt hop3 (0xFB) push data: remaining=${remaining}, decoded=${data.length}`
        );
      }
      sIn.readU8();
      const prev = data[data.length - 2];
      put(prev);
      const val = read255Short(sIn);
      put(val);
      put(prev);
      remaining -= 3;
    } else if (code === 252) {
      if (remaining < 5 || data.length < 2) {
        throw new EotError(
          "CORRUPT_HOPCODE_DATA" /* CorruptHopcodeData */,
          `corrupt hop4 (0xFC) push data: remaining=${remaining}, decoded=${data.length}`
        );
      }
      sIn.readU8();
      const prev = data[data.length - 2];
      put(prev);
      const c = read255Short(sIn);
      put(c);
      put(prev);
      const d = read255Short(sIn);
      put(d);
      put(prev);
      remaining -= 5;
    } else {
      const v = read255Short(sIn);
      put(v);
      remaining -= 1;
    }
  }
  flush();
}
function makeGlyphFlags(x, y, onCurve, firstTime) {
  let flags = 0;
  if (onCurve) {
    flags |= FLG_ON_CURVE;
  }
  if (!firstTime && x === 0) {
    flags |= FLG_X_SAME;
  } else if (x > -256 && x < 0) {
    flags |= FLG_X_SHORT;
  } else if (x >= 0 && x < 256) {
    flags |= FLG_X_SHORT | FLG_X_SAME;
  }
  if (!firstTime && y === 0) {
    flags |= FLG_Y_SAME;
  } else if (y > -256 && y < 0) {
    flags |= FLG_Y_SHORT;
  } else if (y >= 0 && y < 256) {
    flags |= FLG_Y_SHORT | FLG_Y_SAME;
  }
  return flags;
}
function decodeSimpleGlyph(numContours, streams, out, calcBBox, minX, minY, maxX, maxY) {
  if (numContours === 0) {
    return;
  }
  const sGlyph = streams[0];
  out.writeS16(numContours);
  const bboxPos = out.pos;
  if (calcBBox) {
    minX = INT16_MAX;
    minY = INT16_MAX;
    maxX = INT16_MIN;
    maxY = INT16_MIN;
    out.writeS16(0);
    out.writeS16(0);
    out.writeS16(0);
    out.writeS16(0);
  } else {
    out.writeS16(minX);
    out.writeS16(minY);
    out.writeS16(maxX);
    out.writeS16(maxY);
  }
  let totalPoints = 0;
  for (let c = 0; c < numContours; c++) {
    if (c === 0) {
      totalPoints = 1;
    }
    const pointsInContour = read255UShort(sGlyph);
    totalPoints += pointsInContour;
    out.writeU16(totalPoints - 1);
  }
  if (totalPoints > 65535 || totalPoints > sGlyph.size - sGlyph.pos) throw new Error("Font glyph points budget exceeded or truncated");
  const flagBytes = new Uint8Array(boundedLength(totalPoints));
  for (let i = 0; i < totalPoints; i++) {
    flagBytes[i] = sGlyph.readU8();
  }
  const xDeltas = new Int16Array(boundedLength(totalPoints, 2));
  const yDeltas = new Int16Array(boundedLength(totalPoints, 2));
  const onCurve = new Uint8Array(boundedLength(totalPoints));
  let cumulativeX = 0;
  let cumulativeY = 0;
  for (let i = 0; i < totalPoints; i++) {
    const flag = flagBytes[i];
    onCurve[i] = flag & 128 ? 0 : 1;
    const enc = TRIPLET_ENCODINGS[flag & 127];
    let dx = sGlyph.readNBits(enc.xBits) + enc.deltaX;
    let dy = sGlyph.readNBits(enc.yBits) + enc.deltaY;
    if (enc.xSign !== 0) {
      dx *= enc.xSign;
    }
    if (enc.ySign !== 0) {
      dy *= enc.ySign;
    }
    xDeltas[i] = dx;
    yDeltas[i] = dy;
    cumulativeX += xDeltas[i];
    cumulativeY += yDeltas[i];
    if (calcBBox) {
      if (cumulativeX < minX) {
        minX = cumulativeX;
      }
      if (cumulativeX > maxX) {
        maxX = cumulativeX;
      }
      if (cumulativeY < minY) {
        minY = cumulativeY;
      }
      if (cumulativeY > maxY) {
        maxY = cumulativeY;
      }
    }
  }
  const codeSizeLocation = out.pos;
  out.writeU16(0);
  const pushCount = read255UShort(sGlyph);
  decodePushInstructions(streams[1], out, pushCount);
  const codeSize = read255UShort(sGlyph);
  if (codeSize > 0) {
    streams[2].copyTo(out, codeSize);
  }
  const unpackedCodeSize = out.pos - (codeSizeLocation + 2);
  const savedPos = out.pos;
  out.seekAbsolute(codeSizeLocation);
  out.writeU16(unpackedCodeSize);
  out.seekAbsolute(savedPos);
  for (let i = 0; i < totalPoints; i++) {
    const f = makeGlyphFlags(xDeltas[i], yDeltas[i], onCurve[i] !== 0, i === 0);
    out.writeU8(f);
  }
  for (let i = 0; i < totalPoints; i++) {
    const x = xDeltas[i];
    if (i === 0 || x !== 0) {
      const absX = Math.abs(x);
      if (absX < 256) {
        out.writeU8(absX);
      } else {
        out.writeS16(x);
      }
    }
  }
  for (let i = 0; i < totalPoints; i++) {
    const y = yDeltas[i];
    if (i === 0 || y !== 0) {
      const absY = Math.abs(y);
      if (absY < 256) {
        out.writeU8(absY);
      } else {
        out.writeS16(y);
      }
    }
  }
  if (calcBBox) {
    const endPos = out.pos;
    out.seekAbsolute(bboxPos);
    out.writeS16(minX);
    out.writeS16(minY);
    out.writeS16(maxX);
    out.writeS16(maxY);
    out.seekAbsolute(endPos);
  }
}
function decodeCompositeGlyph(streams, out) {
  const sGlyph = streams[0];
  out.writeS16(-1);
  out.writeS16(sGlyph.readS16());
  out.writeS16(sGlyph.readS16());
  out.writeS16(sGlyph.readS16());
  out.writeS16(sGlyph.readS16());
  let flags = 0;
  do {
    flags = sGlyph.readU16();
    const glyphIndex = sGlyph.readU16();
    out.writeU16(flags);
    out.writeU16(glyphIndex);
    let argBytes;
    if (flags & ARG_1_AND_2_ARE_WORDS) {
      argBytes = 4;
    } else {
      argBytes = 2;
    }
    sGlyph.copyTo(out, argBytes);
    let transformBytes = 0;
    if (flags & HAVE_2_BY_2) {
      transformBytes = 8;
    } else if (flags & HAVE_XY_SCALE) {
      transformBytes = 4;
    } else if (flags & HAVE_SCALE) {
      transformBytes = 2;
    }
    if (transformBytes > 0) {
      sGlyph.copyTo(out, transformBytes);
    }
  } while (flags & MORE_COMPONENTS);
  if (flags & HAVE_INSTRUCTIONS) {
    const numInstrPos = out.pos;
    out.writeU16(0);
    const pushCount = read255UShort(sGlyph);
    decodePushInstructions(streams[1], out, pushCount);
    const codeSize = read255UShort(sGlyph);
    if (codeSize > 0) {
      streams[2].copyTo(out, codeSize);
    }
    const numInstr = out.pos - (numInstrPos + 2);
    const savedPos = out.pos;
    out.seekAbsolute(numInstrPos);
    out.writeU16(numInstr);
    out.seekAbsolute(savedPos);
  }
}
function decodeGlyph(streams, out) {
  const numContours = streams[0].readS16();
  if (numContours < 0) {
    decodeCompositeGlyph(streams, out);
  } else if (numContours === 32767) {
    const actualContours = streams[0].readS16();
    const xMin = streams[0].readS16();
    const yMin = streams[0].readS16();
    const xMax = streams[0].readS16();
    const yMax = streams[0].readS16();
    decodeSimpleGlyph(actualContours, streams, out, false, xMin, yMin, xMax, yMax);
  } else {
    decodeSimpleGlyph(numContours, streams, out, true, 0, 0, 0, 0);
  }
}
function populateGlyfAndLoca(glyf, loca, headData, maxpData, streams) {
  const numGlyphs = maxpData.numGlyphs;
  streams[0].seekAbsolute(glyf.offset);
  streams[1].seekAbsolute(0);
  streams[2].seekAbsolute(0);
  const maxGlyphSize = 5 * 2 + // header (numContours + bbox)
  2 * maxpData.maxContours + // endPtsOfContours
  2 + // instructionLength
  maxpData.maxSizeOfInstructions + 256 + // instructions + padding
  5 * maxpData.maxPoints + // flags + coordinates
  4 * maxpData.maxComponentElements * 6 + // composite components
  256;
  const outStream = new Stream(null, 0);
  outStream.reserve(numGlyphs * 256);
  const locaOffsets = [0];
  for (let i = 0; i < numGlyphs; i++) {
    outStream.pos;
    outStream.reserve(outStream.pos + maxGlyphSize);
    decodeGlyph(streams, outStream);
    if (outStream.pos & 1) {
      outStream.writeU8(0);
    }
    locaOffsets.push(outStream.pos);
  }
  const useLongLoca = headData.indexToLocFormat === 1 || outStream.pos > 131070;
  if (useLongLoca && headData.indexToLocFormat === 0) {
    headData.table.buf[50] = 0;
    headData.table.buf[51] = 1;
    headData.indexToLocFormat = 1;
  }
  const locaStream = new Stream(null, 0);
  locaStream.reserve((numGlyphs + 1) * (useLongLoca ? 4 : 2));
  for (const offset of locaOffsets) {
    if (useLongLoca) {
      locaStream.writeU32(offset);
    } else {
      locaStream.writeU16(offset / 2);
    }
  }
  glyf.buf = outStream.toUint8Array();
  glyf.bufSize = glyf.buf.length;
  loca.buf = locaStream.toUint8Array();
  loca.bufSize = loca.buf.length;
}
function parseHead(table) {
  const s = new Stream(table.buf, table.bufSize);
  s.seekAbsolute(50);
  const indexToLocFormat = s.readS16();
  if (indexToLocFormat !== 0 && indexToLocFormat !== 1) {
    throw new EotError(
      "MALFORMED_HEAD_TABLE" /* MalformedHeadTable */,
      `invalid head indexToLocFormat: ${indexToLocFormat}`
    );
  }
  return { indexToLocFormat, table };
}
function parseMaxp(table) {
  const s = new Stream(table.buf, table.bufSize);
  const version = s.readU32();
  const numGlyphs = s.readU16();
  let maxPoints = 0;
  let maxContours = 0;
  let maxSizeOfInstructions = 0;
  let maxComponentElements = 0;
  if (version === 65536) {
    maxPoints = s.readU16();
    maxContours = s.readU16();
    s.readU16();
    s.readU16();
    s.readU16();
    s.readU16();
    s.readU16();
    s.readU16();
    s.readU16();
    s.readU16();
    maxSizeOfInstructions = s.readU16();
    maxComponentElements = s.readU16();
  }
  return {
    numGlyphs,
    maxPoints,
    maxContours,
    maxSizeOfInstructions,
    maxComponentElements
  };
}
function parseCTF(streams, options) {
  const s0 = streams[0];
  s0.readU32();
  const numTables = s0.readU16();
  if (numTables > 4096 || numTables * 16 > s0.size - 12) throw new Error("Font table directory budget exceeded");
  let declaredTableBytes = 12 + numTables * 16;
  s0.readU16();
  s0.readU16();
  s0.readU16();
  const tables = [];
  let glyfIdx = -1;
  let locaIdx = -1;
  let maxpIdx = -1;
  let headIdx = -1;
  let hmtxIdx = -1;
  for (let i = 0; i < numTables; i++) {
    const tag = s0.readChar() + s0.readChar() + s0.readChar() + s0.readChar();
    s0.seekRelative(4);
    const offset = s0.readU32();
    const size = s0.readU32();
    declaredTableBytes += Math.ceil(size / 4) * 4;
    boundedLength(declaredTableBytes);
    const table = {
      tag,
      offset,
      bufSize: size,
      buf: new Uint8Array(boundedLength(0)),
      checksum: 0
    };
    const idx = tables.length;
    tables.push(table);
    if (tag === "glyf") {
      glyfIdx = idx;
    } else if (tag === "loca") {
      locaIdx = idx;
    } else if (tag === "maxp") {
      maxpIdx = idx;
    } else if (tag === "head") {
      headIdx = idx;
    } else if (tag === "hmtx") {
      hmtxIdx = idx;
    } else ;
  }
  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];
    if (table.tag === "glyf" || table.tag === "loca") {
      continue;
    }
    if (table.tag === "cvt ") {
      const previousSize = Math.ceil(table.bufSize / 4) * 4;
      const remaining = officeviewByteLimit - (declaredTableBytes - previousSize);
      withDecoderByteLimit(Math.max(0, remaining), () => unpackCVT(table, s0));
      // Later CVT entries must see expanded predecessors, not their declared sizes.
      declaredTableBytes += Math.ceil(table.bufSize / 4) * 4 - previousSize;
      boundedLength(declaredTableBytes);
      continue;
    }
    s0.seekAbsolute(table.offset);
    if (table.bufSize > s0.size - table.offset) {
      throw new EotError("INSUFFICIENT_BYTES" /* InsufficientBytes */, `truncated ${table.tag} table`);
    }
    const buf = new Uint8Array(boundedLength(table.bufSize));
    for (let b = 0; b < table.bufSize; b++) {
      buf[b] = s0.readU8();
    }
    table.buf = buf;
    if (table.tag === "head") {
      if (table.bufSize < 12) {
        throw new EotError(
          "MALFORMED_HEAD_TABLE" /* MalformedHeadTable */,
          `head table too small: ${table.bufSize} bytes (need at least 12)`
        );
      }
      table.buf[8] = 0;
      table.buf[9] = 0;
      table.buf[10] = 0;
      table.buf[11] = 0;
    }
  }
  if (maxpIdx < 0) {
    throw new EotError("NO_MAXP_TABLE" /* NoMaxpTable */, "CTF font is missing a maxp table");
  }
  if (headIdx < 0) {
    throw new EotError("NO_HEAD_TABLE" /* NoHeadTable */, "CTF font is missing a head table");
  }
  if (hmtxIdx < 0) {
    throw new EotError("NO_HMTX_TABLE" /* NoHmtxTable */, "CTF font is missing an hmtx table");
  }
  const headData = parseHead(tables[headIdx]);
  const maxpData = parseMaxp(tables[maxpIdx]);
  for (const table of tables) {
    if (table.tag === "VDMX") {
      const remaining = officeviewByteLimit - 12 - tables.length * 16 - tables.reduce((sum, other) => sum + (other === table ? 0 : Math.ceil(other.bufSize / 4) * 4), 0);
      table.buf = withDecoderByteLimit(Math.max(0, remaining), () => decodeVdmx(table.buf));
      table.bufSize = table.buf.length;
    } else if (table.tag === "hdmx") {
      const version = table.buf.length >= 2 ? table.buf[0] << 8 | table.buf[1] : -1;
      if (version === 0) {
        const hhea = tables.find((candidate) => candidate.tag === "hhea");
        if (!hhea || hhea.buf.length < 36) {
          throw new EotError("CORRUPT_FILE" /* CorruptFile */, "compressed hdmx requires a complete hhea table");
        }
        const head = tables[headIdx].buf;
        const remaining = officeviewByteLimit - 12 - tables.length * 16 - tables.reduce((sum, other) => sum + (other === table ? 0 : Math.ceil(other.bufSize / 4) * 4), 0);
        table.buf = withDecoderByteLimit(Math.max(0, remaining), () => decodeHdmx(table.buf, {
          numGlyphs: maxpData.numGlyphs,
          unitsPerEm: head[18] << 8 | head[19],
          numberOfHMetrics: hhea.buf[34] << 8 | hhea.buf[35],
          hmtx: tables[hmtxIdx].buf
        }));
      } else {
        table.buf = decodeHdmx(table.buf);
      }
      table.bufSize = table.buf.length;
    }
  }
  if (glyfIdx >= 0) {
    if (locaIdx < 0) {
      const locaTable = {
        tag: "loca",
        offset: 0,
        bufSize: 0,
        buf: new Uint8Array(boundedLength(0)),
        checksum: 0
      };
      locaIdx = tables.length;
      tables.push(locaTable);
    }
    const otherBytes = 12 + tables.length * 16 + tables.reduce((sum, table, i) => sum + (i === glyfIdx || i === locaIdx ? 0 : Math.ceil(table.bufSize / 4) * 4), 0);
    const locaBudget = (maxpData.numGlyphs + 1) * 4;
    withDecoderByteLimit(Math.max(0, officeviewByteLimit - otherBytes - locaBudget), () => populateGlyfAndLoca(tables[glyfIdx], tables[locaIdx], headData, maxpData, streams));
  }
  return { tables };
}

// src/ahuff.ts
function bitsUsed(x) {
  if (x <= 0) {
    return 1;
  }
  return 32 - Math.clz32(x);
}
var AHuff = class _AHuff {
  bio;
  range;
  tree;
  /** Maps symbol value -> current tree index of its leaf node. */
  symbolIndex;
  /** Number of bits that encode a "full-size" symbol (ceil(log2(range))). */
  bitCount;
  /**
   * Secondary bit width used for large-range trees.
   * 0 when range <= 256 (small tree path).
   */
  bitCount2;
  static ROOT = 1;
  constructor(bio, range) {
    this.bio = bio;
    this.range = range;
    this.bitCount = bitsUsed(range - 1);
    this.bitCount2 = 0;
    if (range > 256 && range < 512) {
      this.bitCount2 = bitsUsed(range - 256 - 1) + 1;
    }
    if (!Number.isInteger(range) || range < 1 || range > 512) throw new Error("Invalid Huffman range");
    const treeSize = 2 * range;
    this.tree = Array.from({ length: treeSize });
    for (let i = 0; i < treeSize; i++) {
      this.tree[i] = { up: 0, left: 0, right: 0, code: -1, weight: 0 };
    }
    for (let i = 2; i < treeSize; i++) {
      this.tree[i].up = i >> 1;
      this.tree[i].weight = 1;
    }
    for (let i = 1; i < range; i++) {
      this.tree[i].left = 2 * i;
      this.tree[i].right = 2 * i + 1;
      this.tree[i].code = -1;
    }
    for (let i = 0; i < range; i++) {
      const leafIdx = range + i;
      this.tree[leafIdx].code = i;
      this.tree[leafIdx].left = -1;
      this.tree[leafIdx].right = -1;
    }
    this.symbolIndex = Array.from({ length: range });
    for (let i = 0; i < range; i++) {
      this.symbolIndex[i] = range + i;
    }
    this.initWeight(_AHuff.ROOT);
    if (this.bitCount2 !== 0) {
      this.updateWeight(this.symbolIndex[256]);
      this.updateWeight(this.symbolIndex[257]);
      const dup2Sym = range - 3;
      for (let i = 0; i < 12; i++) {
        this.updateWeight(this.symbolIndex[dup2Sym]);
      }
      const dup4Sym = range - 2;
      for (let i = 0; i < 6; i++) {
        this.updateWeight(this.symbolIndex[dup4Sym]);
      }
    } else {
      for (let j = 0; j < 2; j++) {
        for (let i = 0; i < range; i++) {
          this.updateWeight(this.symbolIndex[i]);
        }
      }
    }
  }
  // --------------------------------------------------------------------
  // Public API
  // --------------------------------------------------------------------
  /**
   * Decode one symbol from the bit stream.
   *
   * Starting at ROOT, read one bit at a time:
   *   - 0 → go left
   *   - 1 → go right
   * Continue until a leaf (code >= 0) is reached.  Then update the
   * tree weights and return the symbol code.
   */
  readSymbol() {
    let a = _AHuff.ROOT;
    let symbol;
    do {
      a = this.bio.inputBit() ? this.tree[a].right : this.tree[a].left;
      symbol = this.tree[a].code;
    } while (symbol < 0);
    this.updateWeight(a);
    return symbol;
  }
  // --------------------------------------------------------------------
  // Private helpers
  // --------------------------------------------------------------------
  /**
   * Increment the weight of node `a` and propagate up to ROOT,
   * swapping nodes as necessary to maintain the sibling property
   * (nodes in non-increasing weight order by index).
   *
   * Algorithm:
   *   For each node from `a` up to (but not including) ROOT:
   *     1. Look at the predecessor (a-1).
   *     2. If it has the same weight, scan backwards to find the first
   *        node with that weight.
   *     3. Swap `a` with that first node (unless it is ROOT or `a`'s
   *        own parent) to restore ordering.
   *     4. Increment `a`'s weight.
   *     5. Move to `a`'s parent.
   *   Finally increment ROOT's weight.
   */
  updateWeight(a) {
    const tree = this.tree;
    for (; a !== _AHuff.ROOT; a = tree[a].up) {
      const weightA = tree[a].weight;
      let b = a - 1;
      if (tree[b].weight === weightA) {
        do {
          b--;
        } while (tree[b].weight === weightA);
        b++;
        if (b > _AHuff.ROOT) {
          this.swapNodes(a, b);
          a = b;
        }
      }
      tree[a].weight = weightA + 1;
    }
    tree[_AHuff.ROOT].weight++;
  }
  /**
   * Swap two nodes in the tree while keeping the parent linkage
   * consistent.
   *
   * What gets swapped: left, right, code, weight — everything that
   * defines the *content* of the node.  The `up` pointer stays with
   * the position (the parent still points here).
   *
   * After the content swap we must:
   *   1. Fix children's `up` pointers (they now live under the other
   *      position).
   *   2. Fix `symbolIndex` for leaves so we can still find them by
   *      symbol value.
   */
  swapNodes(a, b) {
    const tree = this.tree;
    const upa = tree[a].up;
    const upb = tree[b].up;
    const tmp = tree[a];
    tree[a] = tree[b];
    tree[b] = tmp;
    tree[a].up = upa;
    tree[b].up = upb;
    let code = tree[a].code;
    if (code < 0) {
      tree[tree[a].left].up = a;
      tree[tree[a].right].up = a;
    } else {
      this.symbolIndex[code] = a;
    }
    code = tree[b].code;
    if (code < 0) {
      tree[tree[b].left].up = b;
      tree[tree[b].right].up = b;
    } else {
      this.symbolIndex[code] = b;
    }
  }
  /**
   * Recursively compute weights for internal nodes after the initial
   * tree construction.  Leaf weights are already set to 1.
   *
   * weight(internal) = weight(left) + weight(right)
   */
  initWeight(a) {
    const node = this.tree[a];
    if (node.code >= 0) {
      return node.weight;
    }
    node.weight = this.initWeight(node.left) + this.initWeight(node.right);
    return node.weight;
  }
};

// src/bitio.ts
var BitIO = class {
  data;
  index;
  size;
  bitBuffer = 0;
  bitCount = 0;
  /**
   * @param data   Source byte buffer.
   * @param offset Starting byte offset into `data`.
   * @param size   Absolute end index into `data` (exclusive) — reading stops
   *               once `index` reaches it. Defaults to `data.length`. Clamped
   *               to `data.length` so an over-large value cannot read past the
   *               end of the buffer and silently yield zero bits.
   */
  constructor(data, offset = 0, size) {
    this.data = data;
    this.index = offset;
    this.size = Math.min(size ?? data.length, data.length);
  }
  /**
   * Read a single bit from the stream.
   *
   * Mirrors `MTX_BITIO_input_bit`:
   *   - If `bitCount` has reached 0, load the next byte into `bitBuffer`
   *     and reset `bitCount` to 7.
   *   - Shift `bitBuffer` left by 1.
   *   - Return whether bit 8 (0x100) is set (i.e. the MSB that was
   *     shifted out of the original byte value).
   */
  inputBit() {
    if (this.bitCount === 0) {
      if (this.index >= this.size) {
        throw new EotError("INSUFFICIENT_BYTES" /* InsufficientBytes */, "BitIO: end of data");
      }
      this.bitBuffer = this.data[this.index++];
      this.bitCount = 8;
    }
    this.bitCount--;
    this.bitBuffer <<= 1;
    return (this.bitBuffer & 256) !== 0;
  }
  /**
   * Read an unsigned integer of `numberOfBits` width, MSB first.
   *
   * Mirrors `MTX_BITIO_ReadValue`: accumulates bits from the most
   * significant down to the least significant.
   */
  readValue(numberOfBits) {
    let value = 0;
    for (let i = numberOfBits - 1; i >= 0; i--) {
      value <<= 1;
      if (this.inputBit()) {
        value |= 1;
      }
    }
    return value >>> 0;
  }
};

// src/lzcomp.ts
var LEN_WIDTH = 3;
var DIST_WIDTH = 3;
var BIT_RANGE = LEN_WIDTH - 1;
var MAX_2BYTE_DIST = 512;
var PRELOAD_SIZE = 2 * 32 * 96 + 4 * 256;
var LEN_MIN = 2;
var DIST_MIN = 1;
var MAX_OUT_LEN = (1 << 24) - 1;
var MAX_OUT = 32 * 1024 * 1024;
var RLE_INITIAL = 0;
var RLE_NORMAL = 1;
var RLE_SEEN_ESCAPE = 2;
var RLE_NEED_BYTE = 3;
function setDistRange(length) {
  let numDistRanges = 1;
  let distMax = DIST_MIN + ((1 << DIST_WIDTH * numDistRanges) - 1);
  while (distMax < length) {
    numDistRanges++;
    if (numDistRanges > 8) {
      throw new EotError("MTX_ERROR" /* MtxError */, "LZCOMP setDistRange: numDistRanges exceeds bound (8)");
    }
    distMax = DIST_MIN + ((1 << DIST_WIDTH * numDistRanges) - 1);
  }
  const DUP2 = 256 + (1 << LEN_WIDTH) * numDistRanges;
  const DUP4 = DUP2 + 1;
  const DUP6 = DUP4 + 1;
  const NUM_SYMS = DUP6 + 1;
  return { numDistRanges, distMax, DUP2, DUP4, DUP6, NUM_SYMS };
}
function initializeModel(window) {
  let i = 0;
  for (let k = 0; k < 32; k++) {
    for (let j2 = 0; j2 < 96; j2++) {
      window[i++] = k;
      window[i++] = j2;
    }
  }
  let j = 0;
  while (i < PRELOAD_SIZE && j < 256) {
    window[i++] = j;
    window[i++] = j;
    window[i++] = j;
    window[i++] = j;
    j++;
  }
}
function decodeLength(lenEcoder, symbol, numDistRangesOut) {
  const mask = 1 << BIT_RANGE;
  let firstTime = true;
  let value = 0;
  let done;
  let iters = 0;
  do {
    if (++iters > 16) {
      throw new EotError("MTX_ERROR" /* MtxError */, "LZCOMP decodeLength: iteration cap exceeded");
    }
    let bits;
    if (firstTime) {
      bits = symbol - 256;
      firstTime = false;
      numDistRangesOut[0] = Math.floor(bits / (1 << LEN_WIDTH)) + 1;
      bits %= 1 << LEN_WIDTH;
    } else {
      bits = lenEcoder.readSymbol();
    }
    done = (bits & mask) === 0;
    bits &= ~mask;
    value <<= BIT_RANGE;
    value |= bits;
  } while (!done);
  value += LEN_MIN;
  return value;
}
function decodeDistance(distEcoder, numDistRanges) {
  let value = 0;
  for (let i = numDistRanges; i > 0; i--) {
    const bits = distEcoder.readSymbol();
    value <<= DIST_WIDTH;
    value |= bits;
  }
  value += DIST_MIN;
  return value;
}
function lzcompDecompress(data, size, version) {
  const bio = new BitIO(data, 0, size);
  let usingRunLength;
  if (version === 1) {
    usingRunLength = false;
  } else {
    usingRunLength = bio.inputBit();
  }
  const distEcoder = new AHuff(bio, 1 << DIST_WIDTH);
  const lenEcoder = new AHuff(bio, 1 << LEN_WIDTH);
  const outLen = bio.readValue(24);
  if (outLen > MAX_OUT_LEN) {
    throw new EotError("MTX_ERROR" /* MtxError */, `LZCOMP outLen ${outLen} exceeds maximum (${MAX_OUT_LEN})`);
  }
  const { DUP2, DUP4, DUP6, NUM_SYMS } = setDistRange(outLen);
  const symEcoder = new AHuff(bio, NUM_SYMS);
  const windowSize = PRELOAD_SIZE + outLen;
  const win = new Uint8Array(boundedLength(windowSize));
  initializeModel(win);
  const base = PRELOAD_SIZE;
  let outBufSize = outLen;
  let outBuf = new Uint8Array(boundedLength(outBufSize));
  let outIdx = 0;
  let rleState = RLE_INITIAL;
  let rleEscape = 0;
  let rleCount = 0;
  const emitByte = (byte) => {
    if (!usingRunLength) {
      if (outIdx >= outBufSize) {
        outBufSize = Math.max(1, outBufSize + (outBufSize >>> 1));
        if (outBufSize > MAX_OUT) {
          throw new EotError("MTX_ERROR" /* MtxError */, "LZCOMP output exceeds maximum size budget");
        }
        const tmp = new Uint8Array(boundedLength(outBufSize));
        tmp.set(outBuf);
        outBuf = tmp;
      }
      outBuf[outIdx++] = byte;
      return;
    }
    switch (rleState) {
      case RLE_INITIAL:
        rleEscape = byte;
        rleState = RLE_NORMAL;
        break;
      case RLE_NORMAL:
        if (byte === rleEscape) {
          rleState = RLE_SEEN_ESCAPE;
        } else {
          if (outIdx >= outBufSize) {
            outBufSize = Math.max(1, outBufSize + (outBufSize >>> 1));
            if (outBufSize > MAX_OUT) {
              throw new EotError("MTX_ERROR" /* MtxError */, "LZCOMP output exceeds maximum size budget");
            }
            const tmp = new Uint8Array(boundedLength(outBufSize));
            tmp.set(outBuf);
            outBuf = tmp;
          }
          outBuf[outIdx++] = byte;
        }
        break;
      case RLE_SEEN_ESCAPE:
        rleCount = byte;
        if (rleCount === 0) {
          if (outIdx >= outBufSize) {
            outBufSize = Math.max(1, outBufSize + (outBufSize >>> 1));
            if (outBufSize > MAX_OUT) {
              throw new EotError("MTX_ERROR" /* MtxError */, "LZCOMP output exceeds maximum size budget");
            }
            const tmp = new Uint8Array(boundedLength(outBufSize));
            tmp.set(outBuf);
            outBuf = tmp;
          }
          outBuf[outIdx++] = rleEscape;
          rleState = RLE_NORMAL;
        } else {
          rleState = RLE_NEED_BYTE;
        }
        break;
      case RLE_NEED_BYTE: {
        if (outIdx + rleCount > outBufSize) {
          outBufSize = outIdx + rleCount + (outBufSize >>> 1);
          if (outBufSize > MAX_OUT) {
            throw new EotError("MTX_ERROR" /* MtxError */, "LZCOMP output exceeds maximum size budget");
          }
          const tmp = new Uint8Array(boundedLength(outBufSize));
          tmp.set(outBuf);
          outBuf = tmp;
        }
        for (let i = 0; i < rleCount; i++) {
          outBuf[outIdx++] = byte;
        }
        rleState = RLE_NORMAL;
        break;
      }
    }
  };
  let pos = 0;
  for (; pos < outLen; ) {
    const symbol = symEcoder.readSymbol();
    let value;
    if (symbol < 256) {
      value = symbol;
    } else if (symbol === DUP2) {
      value = win[base + pos - 2];
    } else if (symbol === DUP4) {
      value = win[base + pos - 4];
    } else if (symbol === DUP6) {
      value = win[base + pos - 6];
    } else {
      const numDistRangesRef = [0];
      let length = decodeLength(lenEcoder, symbol, numDistRangesRef);
      const distance = decodeDistance(distEcoder, numDistRangesRef[0]);
      if (distance >= MAX_2BYTE_DIST) {
        length++;
      }
      const start = base + pos - distance - length + 1;
      if (length < 1 || length > outLen - pos || start < 0 || start >= base + pos) throw new Error("Invalid LZ copy bounds");
      for (let j = 0; j < length; j++) {
        value = win[start + j];
        win[base + pos] = value;
        pos++;
        emitByte(value);
      }
      continue;
    }
    win[base + pos] = value;
    pos++;
    emitByte(value);
  }
  if (pos !== outLen) {
    throw new EotError(
      "MTX_ERROR" /* MtxError */,
      `LZCOMP decode overran output: pos=${pos}, expected=${outLen}`
    );
  }
  if (usingRunLength && rleState !== RLE_NORMAL && !(outLen === 0 && rleState === RLE_INITIAL)) throw new Error("Truncated LZ run length sequence");
  return outBuf.subarray(0, outIdx);
}

// src/sfnt-builder.ts
function lgFloor(n) {
  let ret = 0;
  while (n > 1) {
    n = Math.floor(n / 2);
    ret++;
  }
  return ret;
}
function maxPow2(n) {
  return 1 << lgFloor(n);
}
function writeOffsetTable(ctr, out) {
  const numTables = ctr.tables.length;
  const searchRange = maxPow2(numTables) * 16;
  const entrySelector = lgFloor(numTables);
  const rangeShift = numTables * 16 - searchRange;
  out.writeU32(65536);
  out.writeU16(numTables);
  out.writeU16(searchRange);
  out.writeU16(entrySelector);
  out.writeU16(rangeShift);
}
function writeTableDirectory(ctr, out) {
  for (const table of [...ctr.tables].sort((a, b) => a.tag < b.tag ? -1 : a.tag > b.tag ? 1 : 0)) {
    out.writeU8(table.tag.charCodeAt(0));
    out.writeU8(table.tag.charCodeAt(1));
    out.writeU8(table.tag.charCodeAt(2));
    out.writeU8(table.tag.charCodeAt(3));
    out.writeU32(table.checksum);
    out.writeU32(table.offset);
    out.writeU32(table.bufSize);
  }
}
function writeTableCheckingSum(table, out) {
  table.offset = out.pos;
  let checksum = 0;
  const data = table.buf;
  const len = table.bufSize;
  const fullWords = Math.floor(len / 4);
  const remainder = len % 4;
  for (let i = 0; i < fullWords; i++) {
    const off = i * 4;
    const isAdjustmentWord = table.tag === "head" && off === 8;
    const word = isAdjustmentWord ? 0 : (data[off] << 24 | data[off + 1] << 16 | data[off + 2] << 8 | data[off + 3]) >>> 0;
    checksum = checksum + word >>> 0;
    out.writeU32(word);
  }
  if (remainder > 0) {
    let word = 0;
    for (let j = 0; j < remainder; j++) {
      word |= data[fullWords * 4 + j] << 24 - j * 8;
    }
    word >>>= 0;
    checksum = checksum + word >>> 0;
    out.writeU32(word);
  }
  table.checksum = checksum;
}
function validateTables(ctr) {
  for (const table of ctr.tables) {
    if (!Number.isInteger(table.bufSize) || table.bufSize < 0 || table.bufSize > table.buf.length) {
      throw new EotError("CORRUPT_FILE" /* CorruptFile */, `table ${table.tag} has an invalid buffer size`);
    }
    if (table.tag === "head" && table.bufSize < 12) {
      throw new EotError("MALFORMED_HEAD_TABLE" /* MalformedHeadTable */, "head table is too short for checksumAdjustment");
    }
  }
}
function getTableDirectorySize(ctr) {
  return 16 * ctr.tables.length;
}
function getRequiredSize(ctr) {
  const offsetTableSize = 12;
  const dirSize = getTableDirectorySize(ctr);
  let tableDataSize = 0;
  for (const table of ctr.tables) {
    tableDataSize += Math.ceil(table.bufSize / 4) * 4;
  }
  return offsetTableSize + dirSize + tableDataSize;
}
function dumpContainer(ctr) {
  validateTables(ctr);
  const requiredSize = getRequiredSize(ctr);
  const out = new Stream(new Uint8Array(boundedLength(requiredSize)), 0);
  writeOffsetTable(ctr, out);
  const dirOffset = out.pos;
  const dirSize = getTableDirectorySize(ctr);
  out.pos += dirSize;
  let totalChecksum = 0;
  for (const table of ctr.tables) {
    writeTableCheckingSum(table, out);
    totalChecksum = totalChecksum + table.checksum >>> 0;
  }
  let headTable;
  for (const table of ctr.tables) {
    if (table.tag === "head") {
      headTable = table;
      break;
    }
  }
  const afterTables = out.pos;
  out.pos = dirOffset;
  writeTableDirectory(ctr, out);
  const beginningLen = 12 + dirSize;
  let beginningChecksum = 0;
  const buf = out.buf;
  const beginningWords = Math.floor(beginningLen / 4);
  for (let i = 0; i < beginningWords; i++) {
    const off = i * 4;
    const word = (buf[off] << 24 | buf[off + 1] << 16 | buf[off + 2] << 8 | buf[off + 3]) >>> 0;
    beginningChecksum = beginningChecksum + word >>> 0;
  }
  totalChecksum = totalChecksum + beginningChecksum >>> 0;
  if (!headTable) {
    throw new EotError(
      "NO_HEAD_TABLE" /* NoHeadTable */,
      "cannot assemble SFNT: container is missing a head table"
    );
  }
  const finalChecksum = 2981146554 - totalChecksum >>> 0;
  const adjOffset = headTable.offset + 8;
  buf[adjOffset] = finalChecksum >>> 24 & 255;
  buf[adjOffset + 1] = finalChecksum >>> 16 & 255;
  buf[adjOffset + 2] = finalChecksum >>> 8 & 255;
  buf[adjOffset + 3] = finalChecksum & 255;
  out.pos = afterTables;
  return out.buf.subarray(0, out.pos);
}

// src/mtx-decompress.ts
var ENCRYPTION_KEY = 80;
function unpackMtx(data, size) {
  if (data.length < 10) {
    throw new EotError("INSUFFICIENT_BYTES" /* InsufficientBytes */, "MTX data too small: header requires at least 10 bytes");
  }
  if (!Number.isInteger(size)) {
    throw new EotError(
      "MTX_ERROR" /* MtxError */,
      `MTX data size must be an integer: size=${size}`
    );
  }
  if (size < 10) {
    throw new EotError("INSUFFICIENT_BYTES" /* InsufficientBytes */, "MTX data too small: header requires at least 10 bytes");
  }
  if (size > data.length) {
    throw new EotError(
      "INSUFFICIENT_BYTES" /* InsufficientBytes */,
      `MTX data size exceeds the input buffer: size=${size}, data.length=${data.length}`
    );
  }
  const versionMagic = data[0];
  const offset2 = data[4] << 16 | data[5] << 8 | data[6];
  const offset3 = data[7] << 16 | data[8] << 8 | data[9];
  if (offset2 < 10 || offset3 < offset2 || offset3 > size) {
    throw new EotError(
      "MTX_ERROR" /* MtxError */,
      `MTX header offsets out of bounds: offset2=${offset2}, offset3=${offset3}, size=${size}`
    );
  }
  const offsets = [10, offset2, offset3];
  const blockSizes = [
    Math.max(0, offset2 - 10),
    Math.max(0, offset3 - offset2),
    Math.max(0, size - offset3)
  ];
  const streams = [];
  const decompressedSizes = [];
  for (let i = 0; i < 3; i++) {
    const block = data.subarray(offsets[i], offsets[i] + blockSizes[i]);
    const remaining = officeviewByteLimit - decompressedSizes.reduce((sum, n) => sum + n, 0);
    const decompressed = withDecoderByteLimit(remaining, () => lzcompDecompress(block, blockSizes[i], versionMagic));
    boundedLength(decompressed.length + decompressedSizes.reduce((sum, n) => sum + n, 0));
    streams.push(decompressed);
    decompressedSizes.push(decompressed.length);
  }
  return { streams, sizes: decompressedSizes };
}
function decompressMtx(fontData, options) {
  const encrypted = options?.encrypted ?? false;
  const compressed = options?.compressed ?? true;
  let data;
  if (encrypted) {
    data = new Uint8Array(boundedLength(fontData.length));
    for (let i = 0; i < fontData.length; i++) {
      data[i] = fontData[i] ^ ENCRYPTION_KEY;
    }
  } else {
    data = fontData;
  }
  if (!compressed) {
    return encrypted ? data : data.slice();
  }
  const { streams } = unpackMtx(data, data.length);
  const streamObjects = streams.map((buf) => new Stream(buf, buf.length));
  const container = parseCTF(streamObjects, { onWarn: options?.onWarn });
  return dumpContainer(container);
}
function decompressEotFont(fontData, compressed, encrypted) {
  return decompressMtx(fontData, { encrypted, compressed });
}

// src/eot.ts
var TTEMBED_SUBSET = 1;
var TTEMBED_TTCOMPRESSED = 4;
var TTEMBED_XORENCRYPTDATA = 268435456;
var EOT_MAGIC = 20556;
var EDITING_MASK = 8;
function readU16LE(bytes, at) {
  return bytes[at] | bytes[at + 1] << 8;
}
function readU32LE(bytes, at) {
  return (bytes[at] | bytes[at + 1] << 8 | bytes[at + 2] << 16 | bytes[at + 3] << 24) >>> 0;
}
function decodeUtf16LE(bytes, at, byteLength) {
  let out = "";
  for (let i = 0; i < byteLength; i += 2) {
    out += String.fromCharCode(readU16LE(bytes, at + i));
  }
  return out;
}
var Scanner = class {
  constructor(bytes, start, limit) {
    this.bytes = bytes;
    this.limit = limit;
    this.pos = start;
  }
  bytes;
  limit;
  pos;
  ensure(n) {
    if (this.pos + n > this.limit) {
      throw new EotError(
        "INSUFFICIENT_BYTES" /* InsufficientBytes */,
        `EOT header truncated: need ${n} more byte(s) at offset ${this.pos}`
      );
    }
  }
  u16() {
    this.ensure(2);
    const v = readU16LE(this.bytes, this.pos);
    this.pos += 2;
    return v;
  }
  u32() {
    this.ensure(4);
    const v = readU32LE(this.bytes, this.pos);
    this.pos += 4;
    return v;
  }
  u8() {
    this.ensure(1);
    return this.bytes[this.pos++];
  }
  take(n) {
    this.ensure(n);
    const slice = this.bytes.subarray(this.pos, this.pos + n);
    this.pos += n;
    return slice;
  }
  skip(n) {
    this.ensure(n);
    this.pos += n;
  }
  /** Read a length-prefixed (U16LE byte count) UTF-16LE string. */
  string() {
    this.ensure(2);
    const size = readU16LE(this.bytes, this.pos);
    this.pos += 2;
    if (size % 2 !== 0) {
      throw new EotError(
        "BOGUS_STRING_SIZE" /* BogusStringSize */,
        `EOT string size ${size} is not a multiple of 2 (UTF-16)`
      );
    }
    if (size === 0) {
      return "";
    }
    this.ensure(size);
    const s = decodeUtf16LE(this.bytes, this.pos, size);
    this.pos += size;
    return s;
  }
  /** Read a length-prefixed (U32LE byte count) raw byte array. */
  byteArray() {
    this.ensure(4);
    const size = readU32LE(this.bytes, this.pos);
    this.pos += 4;
    if (size === 0) {
      return new Uint8Array(boundedLength(0));
    }
    return this.take(size);
  }
};
var HeaderTooBig = class extends Error {
};
function parseBody(bytes, version, totalSize, fontDataSize) {
  const HEADER_START = 12;
  const limit = bytes.length - fontDataSize;
  const sc = new Scanner(bytes, HEADER_START, limit);
  const flags = sc.u32();
  const panose = sc.take(10).slice();
  const charset = sc.u8();
  const italic = sc.u8() !== 0;
  const weight = sc.u32();
  const permissions = sc.u16();
  if (sc.u16() !== EOT_MAGIC) {
    throw new EotError("CORRUPT_FILE" /* CorruptFile */, "EOT magic number (0x504C) mismatch");
  }
  const unicodeRange = [sc.u32(), sc.u32(), sc.u32(), sc.u32()];
  const codePageRange = [sc.u32(), sc.u32()];
  const checkSumAdjustment = sc.u32();
  sc.skip(18);
  const familyName = sc.string();
  sc.skip(2);
  const styleName = sc.string();
  sc.skip(2);
  const versionName = sc.string();
  sc.skip(2);
  const fullName = sc.string();
  let rootString = "";
  if (version > 1) {
    sc.skip(2);
    rootString = sc.string();
    if (version === 3) {
      sc.u32();
      sc.u32();
      sc.skip(2);
      const signatureSize = sc.u16();
      sc.skip(signatureSize);
      sc.u32();
      sc.byteArray();
    }
  }
  const fontDataOffset = sc.pos;
  const expectedHeaderSize = totalSize - fontDataSize;
  if (fontDataOffset < expectedHeaderSize) {
    throw new HeaderTooBig();
  }
  return {
    version,
    flags,
    panose,
    charset,
    italic,
    weight,
    permissions,
    unicodeRange,
    codePageRange,
    checkSumAdjustment,
    familyName,
    styleName,
    versionName,
    fullName,
    rootString,
    totalSize,
    fontDataSize,
    fontDataOffset
  };
}
var VERSION_MAGIC = {
  65536: 1,
  131073: 2,
  131074: 3
};
function parseEotMetadata(bytes) {
  if (bytes.length < 8) {
    throw new EotError("INSUFFICIENT_BYTES" /* InsufficientBytes */, "EOT file too small (need at least 8 bytes)");
  }
  const totalSize = readU32LE(bytes, 0);
  const fontDataSize = readU32LE(bytes, 4);
  const metadataLength = totalSize - fontDataSize;
  if (bytes.length < metadataLength) {
    throw new EotError(
      "INSUFFICIENT_BYTES" /* InsufficientBytes */,
      `EOT file shorter than its declared metadata length (${metadataLength})`
    );
  }
  if (bytes.length < 12) {
    throw new EotError("INSUFFICIENT_BYTES" /* InsufficientBytes */, "EOT file too small for a version field");
  }
  const versionMagic = readU32LE(bytes, 8);
  const codedVersion = VERSION_MAGIC[versionMagic];
  if (codedVersion === void 0) {
    throw new EotError(
      "CORRUPT_FILE" /* CorruptFile */,
      `unrecognized EOT version magic 0x${versionMagic.toString(16)}`
    );
  }
  if (12 + fontDataSize > bytes.length) {
    throw new EotError("CORRUPT_FILE" /* CorruptFile */, "EOT font data extends past end of file");
  }
  let tryVersion = codedVersion;
  let bumpedUp = false;
  let knockedDown = false;
  while (true) {
    try {
      const body = parseBody(bytes, tryVersion, totalSize, fontDataSize);
      const flags = body.flags;
      return {
        ...body,
        compressed: (flags & TTEMBED_TTCOMPRESSED) !== 0,
        encrypted: (flags & TTEMBED_XORENCRYPTDATA) !== 0,
        badVersion: tryVersion !== codedVersion
      };
    } catch (err) {
      if (err instanceof HeaderTooBig) {
        if (knockedDown || tryVersion === 3) {
          throw new EotError("CORRUPT_FILE" /* CorruptFile */, "EOT header inconsistent across all versions");
        }
        knockedDown = false;
        bumpedUp = true;
        tryVersion = tryVersion + 1;
        continue;
      }
      if (err instanceof EotError && err.code === "INSUFFICIENT_BYTES" /* InsufficientBytes */) {
        if (bumpedUp || tryVersion === 1) {
          throw new EotError("CORRUPT_FILE" /* CorruptFile */, "EOT header inconsistent across all versions");
        }
        knockedDown = true;
        bumpedUp = false;
        tryVersion = tryVersion - 1;
        continue;
      }
      throw err;
    }
  }
}
function inspectEotProtection(bytes) {
  const metadata = parseEotMetadata(bytes);
  return {
    encryption: metadata.encrypted ? "xor-0x50" : "none",
    passwordProtection: "not_supported_by_eot",
    embeddingPermissions: metadata.permissions,
    rootString: metadata.rootString
  };
}
function eotToTtf(bytes, options) {
  const meta = parseEotMetadata(bytes);
  if (meta.badVersion) {
    options?.onWarn?.(`EOT header version disagrees with its layout; decoded as version ${meta.version}`);
  }
  const fontData = bytes.subarray(meta.fontDataOffset, meta.fontDataOffset + meta.fontDataSize);
  return decompressMtx(fontData, {
    compressed: meta.compressed,
    encrypted: meta.encrypted,
    onWarn: options?.onWarn
  });
}
function canLegallyEdit(metadata) {
  return metadata.permissions === 0 || (metadata.permissions & EDITING_MASK) !== 0;
}

export { EOT_WARN, EotError, EotErrorCode, TTEMBED_SUBSET, TTEMBED_TTCOMPRESSED, TTEMBED_XORENCRYPTDATA, canLegallyEdit, decompressEotFont, decompressMtx, eotToTtf, inspectEotProtection, parseCTF, parseEotMetadata, unpackMtx };
