"""
DM3 binary file writer (template-based)
========================================
Writes Gatan DigitalMicrograph DM3 files (.dm3) from numpy arrays
with pixel calibration (scale, offset).

Uses a real GMS-generated DM3 as a structural template to guarantee
binary compatibility with CrysTBox RingGUI, ImageJ, MATLAB ReadDMFile,
and other strict DM3 readers.

The template is reference_template.dm3 (from the rsciio test suite,
originally produced by GMS). Only the pixel data, data dimensions,
data type codes, and calibration values are surgically patched — all
structural bytes come verbatim from the GMS-produced reference.

This is intentionally DM3-only.  For DM4 use a dedicated tool.
"""

import struct
import numpy as np
import os

# ============ Locate the reference template ============
_REF_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         'reference_template.dm3')

# ============ Verified byte offsets in reference DM3 ============
# All offsets were verified by a full recursive parse of the tag tree.
#
# Header: file_size field at offset 4 (BE uint32).
#   Reference: field = 192688, file = 192708.  Formula: field = total - 20.
#
# Main-image calibrations (ImageList entry 2, all float32 LE):
_CAL_DIM1_ORIGIN = 152441   # float32 LE  (x-axis origin, ref = -786.0)
_CAL_DIM1_SCALE  = 152465   # float32 LE  (x-axis scale, ref = 0.17443286)
_CAL_DIM2_ORIGIN = 152535   # float32 LE  (y-axis origin, ref = -756.0)
_CAL_DIM2_SCALE  = 152559   # float32 LE  (y-axis scale, ref = 0.17443286)
# (Units tags are UTF-16 arrays - kept as-is from the template.)
#
# Data-array info (the %%%% header right before pixel bytes):
_DATA_ELEM_TYPE  = 152656   # BE uint32, DM array element type (ref = 3 = int32)
_DATA_ELEM_COUNT = 152660   # BE uint32, number of elements   (ref = 7569)
_DATA_BODY       = 152664   # first byte of pixel data
_ORIG_DATA_BYTES = 7569 * 4 # 30276 bytes of pixel data in the reference
_SUFFIX_START    = _DATA_BODY + _ORIG_DATA_BYTES  # 182940
#
# Suffix-relative offsets (added to new suffix start address):
_REL_DATATYPE    = 182963 - _SUFFIX_START   #  23
_REL_DIM_NX      = 183001 - _SUFFIX_START   #  61
_REL_DIM_NY      = 183020 - _SUFFIX_START   #  80
_REL_PIXELDEPTH  = 183049 - _SUFFIX_START   # 109

# ============ DM3 type-code tables ============
_ARRAY_TYPES = {
    np.dtype('int8'):    10, np.dtype('uint8'):   10,
    np.dtype('int16'):   2,  np.dtype('uint16'):  4,
    np.dtype('int32'):   3,  np.dtype('uint32'):  5,
    np.dtype('float32'): 6,  np.dtype('float64'): 7,
}
_IMAGE_DTYPES = {
    np.dtype('int8'):    9,  np.dtype('uint8'):   6,
    np.dtype('int16'):   1,  np.dtype('uint16'):  10,
    np.dtype('int32'):   7,  np.dtype('uint32'):  11,
    np.dtype('float32'): 2,  np.dtype('float64'): 12,
}


def write_dm(filepath, data, pixel_scales=None, pixel_units=None,
             pixel_offsets=None, title="", version=3):
    """Write a 2-D array as a CrysTBox-compatible DM3 file.

    Parameters
    ----------
    filepath : str
        Output .dm3 file path.
    data : numpy.ndarray
        2-D image array with shape (height, width).
    pixel_scales : tuple of float, optional
        (y_scale, x_scale).  Default (1.0, 1.0).
    pixel_units : tuple of str, optional
        Ignored - the template keeps "1/nm" from the reference.
    pixel_offsets : tuple of float, optional
        (y_offset, x_offset).  Default (0.0, 0.0).
    title : str, optional
        Ignored - the template keeps the reference image name.
    version : int
        Must be 3.
    """
    if version != 3:
        raise ValueError("Only DM3 (version=3) is supported")
    if data.ndim != 2:
        raise ValueError(f"Expected 2-D array, got {data.ndim}-D")

    # ---- coerce to a supported 4-byte type ----
    if data.dtype == np.float64:
        data = data.astype(np.float32)
    elif data.dtype == np.int64:
        data = data.astype(np.int32)
    elif data.dtype == np.uint64:
        data = data.astype(np.uint32)
    elif data.dtype not in _ARRAY_TYPES:
        data = data.astype(np.float32)
    data = np.ascontiguousarray(data)

    ny, nx = data.shape
    elem_type   = _ARRAY_TYPES[data.dtype]
    image_dtype = _IMAGE_DTYPES[data.dtype]
    pixel_depth = data.dtype.itemsize

    if pixel_scales is None:
        pixel_scales = (1.0, 1.0)
    if pixel_offsets is None:
        pixel_offsets = (0.0, 0.0)

    # ---- load the GMS template ----
    with open(_REF_PATH, 'rb') as f:
        ref = bytearray(f.read())

    prefix = bytearray(ref[:_DATA_BODY])          # everything before pixels
    suffix = bytearray(ref[_SUFFIX_START:])        # everything after pixels

    # ---- patch prefix: calibrations (float32 LE) ----
    # DIM1 = x-axis (columns), DIM2 = y-axis (rows)
    #
    # rsciio convention:  physical = offset + pixel * scale
    # DM3 convention:     physical = (pixel - origin) * scale
    # Therefore:          origin = -offset / scale
    struct.pack_into('<f', prefix, _CAL_DIM1_SCALE,  float(pixel_scales[1]))
    struct.pack_into('<f', prefix, _CAL_DIM2_SCALE,  float(pixel_scales[0]))

    x_scale = float(pixel_scales[1]) if pixel_scales[1] != 0 else 1.0
    y_scale = float(pixel_scales[0]) if pixel_scales[0] != 0 else 1.0
    x_origin = -float(pixel_offsets[1]) / x_scale
    y_origin = -float(pixel_offsets[0]) / y_scale
    struct.pack_into('<f', prefix, _CAL_DIM1_ORIGIN, x_origin)
    struct.pack_into('<f', prefix, _CAL_DIM2_ORIGIN, y_origin)

    # ---- patch prefix: data-array info (BE uint32) ----
    struct.pack_into('>I', prefix, _DATA_ELEM_TYPE,  elem_type)
    struct.pack_into('>I', prefix, _DATA_ELEM_COUNT, nx * ny)

    # ---- pixel data bytes ----
    data_bytes = data.tobytes()

    # ---- patch suffix (LE uint32) ----
    struct.pack_into('<I', suffix, _REL_DATATYPE,   image_dtype)
    struct.pack_into('<I', suffix, _REL_DIM_NX,     nx)
    struct.pack_into('<I', suffix, _REL_DIM_NY,     ny)
    struct.pack_into('<I', suffix, _REL_PIXELDEPTH, pixel_depth)

    # ---- assemble and patch header file-size field ----
    result = bytearray(prefix) + data_bytes + bytearray(suffix)
    struct.pack_into('>I', result, 4, len(result) - 20)

    with open(filepath, 'wb') as f:
        f.write(result)
