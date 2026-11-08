# TESS
## TESS
TESS is an Adafruit Feather compatible embedded Linux board based on the Allwinner T113-S3 SoC with 128MB of DDR3 RAM built-in. We did add some pins that are not in the original spec, these allow for USB and debug UART to be available on regular pinheaders. This SoC does however only have 1 ADC pin availble that is even limited to 1.8V, so the extra analog pins are replaced by digital pins with various mux options. Allowing for multiple I²C busses, UARTs and CAN interfaces to be used.
Some key features of TESS REV 00 are:
- [Allwinner T113-S3 SoC](https://linux-sunxi.org/T113-s3) with dual core Cortex A7 running at 1.2GHz and 128 MB DDR3 RAM
- 512 MB SD-storage soldered down
- [ESP32-C3-MINI-1](https://www.espressif.com/en/module/esp32-c3-mini-1-en) Wi-Fi and Bluetooth module
- battery charger (up to 1.2A selectable in software)
- Microphone
- 3 user LEDs
- [RPI Touch Display 2](https://www.raspberrypi.com/products/touch-display-2/) compatible DSI connector

## Software
Follow the development of our [buildroot Linux.](https://github.com/phyx-be/buildroot-tess)

## Blockdiagram
TESS is designed to be a breakout board for the Allwinner T113-S3 with limited features on the board but maximal flexibility.
![TESS 00 Blockdiagram](media/TESS_Block.png)

## Power
TESS has a 3.3V and 0.9V DC-DC regulator onboard whilst the T113-S3 has 2 LDO's built-in for the 1.8V and 1.5V rails.
![TESS 00 pinout](media/TESS_Power.png)

## Pinout
![TESS 00 pinout](media/TESS_Pinout.png)

## Elecrow PCBA quotation package (whatnick fork)

The upstream Rev 00 exports are preserved. A ten-board **RFQ-only** package is
in [`manufacturing/elecrow-rev00`](manufacturing/elecrow-rev00).

Upload `TESS-rev00-Elecrow-BOM.xlsx` as BOM,
`TESS-rev00-Elecrow-Gerbers.zip` as Gerbers, and
`TESS-rev00-Elecrow-Assembly.zip` as Other Files. BOM quantities are **per board**;
request **10 assemblies**. The assembly archive includes both-side positions,
upstream board drawings, schematic sheets, 3D overviews, paste layers and notes.

**Do not authorise manufacture yet.** Read `PCBA-NOTES.txt` and `manifest.json`
for sourcing conflicts, castellated-edge requirements, uncertain stackup and
thickness, and first-article/test holds. The reported upstream thickness is
about 0.743 mm, not a confirmed fabrication specification. Native CAD sources
and ERC/DRC reports are not present; no clean design-check or DFM claim is made.
No supplier upload, quote submission, purchase or payment has been performed.

Regenerate and validate with Python 3.11 or later:

```powershell
python -m pip install -r requirements-manufacturing.txt
python -m unittest discover -s tests -v
python scripts\prepare_elecrow.py --boards 10
```

Copper, mask and silk are copied unchanged. The sole outline retains the ODB
analytic arcs; separate metric PTH/NPTH drills retain ODB hole centres and
diameters. Source/output hashes and archive-member hashes accompany the package.

