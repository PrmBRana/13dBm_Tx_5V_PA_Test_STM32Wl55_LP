raw_hex_40 = """
76 F0 EB 0D DA 20 10 6E 71 AF 5D 55 04 03 BD 64
77 94 78 36 83 E6 8C 46 DA B3 3B 97 D9 4B 85 E0
61 0B DD C5 CF 43 ED 97
"""
raw_bytes = bytes.fromhex(raw_hex_40.replace('\n', ' '))
print(f"Bytes count: {len(raw_bytes)}")

AX25_FLAG = 0x7E
G3RUH_POLY_TAP1 = 11
G3RUH_POLY_TAP2 = 16
G3RUH_REGISTER_MASK = 0x1FFFF
G3RUH_INITIAL_STATE = 0

sr = G3RUH_INITIAL_STATE
lastNrzi = 1
flagShift = 0
inFrame = False
frameBits = []

def unstuff_and_pack(bits):
    destuffed = []
    ones = 0
    i = 0
    while i < len(bits):
        b = bits[i]
        if b == 1:
            ones += 1
            destuffed.append(1)
            if ones == 5:
                if i + 1 < len(bits):
                    if bits[i+1] == 0:
                        i += 1
                ones = 0
        else:
            ones = 0
            destuffed.append(0)
        i += 1
    
    bytes_out = []
    for j in range(0, len(destuffed) - 7, 8):
        byte_val = 0
        for b in range(8):
            byte_val |= (destuffed[j + b] << b)
        bytes_out.append(byte_val)
    return bytes_out, len(destuffed) % 8

for byteIndex, value in enumerate(raw_bytes):
    for b in range(7, -1, -1):
        scrambledBit = (value >> b) & 1
        feedback = ((sr >> G3RUH_POLY_TAP1) ^ (sr >> G3RUH_POLY_TAP2)) & 1
        nrziBit = scrambledBit ^ feedback
        sr = ((sr << 1) | scrambledBit) & G3RUH_REGISTER_MASK
        dataBit = 1 if (nrziBit == lastNrzi) else 0
        lastNrzi = nrziBit
        
        flagShift = ((flagShift >> 1) | (dataBit << 7)) & 0xFF
        
        if flagShift == AX25_FLAG:
            if not inFrame:
                inFrame = True
                frameBits = []
            else:
                payload_bits = frameBits[:-7]
                bytes_out, rem = unstuff_and_pack(payload_bits)
                print(f"Candidate: {len(bytes_out)} bytes (rem: {rem}): {bytes(bytes_out).hex()}")
                frameBits = []
        else:
            if inFrame:
                frameBits.append(dataBit)

print("In frame at end?", inFrame, "Frame bits left:", len(frameBits))
