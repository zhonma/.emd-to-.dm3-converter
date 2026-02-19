# .emd-to-.dm3-converter
Metadata-preserving TEM file converter between .emd and .dm3 formats. Independent research tool, not affiliated with any vendor.

Transmission Electron Microscopy (TEM) workflows are often fragmented across proprietary file formats. Some analysis tools support .emd, while others rely on .dm3, creating unnecessary friction when transferring data between platforms.

This project provides a vibe-coded research-oriented converter between .emd and .dm3 formats, with a strong focus on preserving embedded metadata during the conversion process. Experimental parameters, acquisition settings, calibration data, and structural metadata are retained to the greatest extent technically possible.

Important Notes

This software is developed independently for research and educational use only.

It is not affiliated with, endorsed by, or supported by any commercial TEM hardware or software manufacturer.

All trademarks and file format names belong to their respective owners.

Users are responsible for ensuring compliance with any applicable software licenses or data usage agreements.

Status

This is an experimental tool. Validation across different instrument generations and software versions is ongoing. Always verify converted files before use in critical workflows.

For usage:
Install all required library.
DO NOT DELETE THE reference_template.dm3!!!
RUN THE CONVERTER BY RUNNING emd_to_dm_converter.py, the other script will automatically run.
