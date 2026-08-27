"""Shared test fixtures."""

import struct

import pytest

# TIFF field type codes used below.
_ASCII = 2
_BYTE = 1
_LONG = 4
_RATIONAL = 5

# ORF is a TIFF variant: same byte layout, but the version word in the header
# is "RO" rather than 42. Building one here means the ORF path can be tested
# without committing a 20 MB photograph to the repository.
ORF_MAGIC = 0x4F52
TIFF_MAGIC = 0x002A


def _entry(tag: int, field_type: int, count: int, payload: bytes, heap: bytearray, base: int):
    """Build one 12-byte IFD entry, spilling oversized values onto the heap."""
    if len(payload) <= 4:
        value = payload.ljust(4, b"\x00")
    else:
        value = struct.pack("<I", base + len(heap))
        heap.extend(payload)
    return struct.pack("<HHI", tag, field_type, count) + value


def _rational(num: int, den: int = 1) -> bytes:
    return struct.pack("<II", num, den)


def _dms(degrees: int, minutes: int, seconds: float) -> bytes:
    """Encode degrees/minutes/seconds as the three rationals EXIF expects."""
    return _rational(degrees) + _rational(minutes) + _rational(int(round(seconds * 100)), 100)


def build_tiff_exif(
    *,
    magic: int = ORF_MAGIC,
    date_taken: str = "2026:08:20 14:19:27",
    latitude=(60, 23, 47.26),
    lat_ref: str = "N",
    longitude=(5, 21, 11.32),
    lon_ref: str = "E",
    altitude_m: float = 312.2,
    altitude_ref: int = 0,
    model: str = "TG-7",
) -> bytes:
    """Build a minimal TIFF/ORF file carrying the tags a field photo would.

    Args:
        magic: Header version word; ORF_MAGIC or TIFF_MAGIC
        date_taken: DateTimeOriginal, in EXIF's own format
        latitude: Degrees, minutes, seconds; None to omit GPS entirely
        lat_ref: "N" or "S"
        longitude: Degrees, minutes, seconds
        lon_ref: "E" or "W"
        altitude_m: Altitude in metres
        altitude_ref: 0 above sea level, 1 below
        model: Camera model string

    Returns:
        The file's bytes, typically under a kilobyte
    """
    has_gps = latitude is not None and longitude is not None

    # Three IFDs laid out back to back, then a heap for values over 4 bytes.
    ifd0_offset = 8
    ifd0_size = 2 + 3 * 12 + 4
    exif_offset = ifd0_offset + ifd0_size
    exif_size = 2 + 1 * 12 + 4
    gps_offset = exif_offset + exif_size
    gps_size = 2 + (6 if has_gps else 0) * 12 + 4
    heap_base = gps_offset + gps_size

    heap = bytearray()

    model_bytes = model.encode("ascii") + b"\x00"
    ifd0 = struct.pack("<H", 3)
    ifd0 += _entry(0x0110, _ASCII, len(model_bytes), model_bytes, heap, heap_base)
    ifd0 += _entry(0x8769, _LONG, 1, struct.pack("<I", exif_offset), heap, heap_base)
    ifd0 += _entry(0x8825, _LONG, 1, struct.pack("<I", gps_offset), heap, heap_base)
    ifd0 += struct.pack("<I", 0)

    date_bytes = date_taken.encode("ascii") + b"\x00"
    exif_ifd = struct.pack("<H", 1)
    exif_ifd += _entry(0x9003, _ASCII, len(date_bytes), date_bytes, heap, heap_base)
    exif_ifd += struct.pack("<I", 0)

    gps_ifd = struct.pack("<H", 6 if has_gps else 0)
    if has_gps:
        gps_ifd += _entry(0x0001, _ASCII, 2, lat_ref.encode() + b"\x00", heap, heap_base)
        gps_ifd += _entry(0x0002, _RATIONAL, 3, _dms(*latitude), heap, heap_base)
        gps_ifd += _entry(0x0003, _ASCII, 2, lon_ref.encode() + b"\x00", heap, heap_base)
        gps_ifd += _entry(0x0004, _RATIONAL, 3, _dms(*longitude), heap, heap_base)
        gps_ifd += _entry(0x0005, _BYTE, 1, bytes([altitude_ref]), heap, heap_base)
        gps_ifd += _entry(
            0x0006, _RATIONAL, 1, _rational(int(round(altitude_m * 10)), 10), heap, heap_base
        )
    gps_ifd += struct.pack("<I", 0)

    header = b"II" + struct.pack("<HI", magic, ifd0_offset)
    return header + ifd0 + exif_ifd + gps_ifd + bytes(heap)


@pytest.fixture
def exif_builder():
    """The TIFF/ORF builder, for tests that need custom tags."""
    return build_tiff_exif


@pytest.fixture
def orf_magic():
    """The ORF header version word."""
    return ORF_MAGIC


@pytest.fixture
def tiff_magic():
    """The standard TIFF header version word."""
    return TIFF_MAGIC


@pytest.fixture
def tiny_orf(tmp_path):
    """A ~250-byte file with ORF magic carrying a date and GPS coordinates."""
    path = tmp_path / "20260820_P8203037.ORF"
    path.write_bytes(build_tiff_exif())
    return path
