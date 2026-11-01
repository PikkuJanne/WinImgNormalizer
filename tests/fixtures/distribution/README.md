# Distribution regression fixtures

The three tiny synthetic files and `expected.json` were supplied by the project
owner for the authorized 1.0.2 launcher repair and automated verification. They
contain no personal media. They are redistributed here under this repository's
MIT license with the owner's request to use them as release regression fixtures.
The original request archive is a test input, not application instructions.

`expected.json` pins every original byte and the independent output assertions:
96 by 64 JPEG, 80 by 48 JPEG, white transparent corner, unchanged MP4 and exit 0.
The test uses Windows GDI+ to decode the resulting JPEGs independently of the
application's ImageMagick engine. No development dependency becomes an application
requirement.
