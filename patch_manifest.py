#!/usr/bin/env python3
"""Patch AndroidManifest.xml to add FileProvider and queries tag for camera support."""
import sys
import re


PROVIDER_XML = '''
    <provider
        android:name="androidx.core.content.FileProvider"
        android:authorities="${applicationId}.fileprovider"
        android:exported="false"
        android:grantUriPermissions="true">
        <meta-data
            android:name="android.support.FILE_PROVIDER_PATHS"
            android:resource="@xml/file_paths" />
    </provider>
'''

QUERIES_XML = '''
    <queries>
        <intent>
            <action android:name="android.media.action.IMAGE_CAPTURE" />
        </intent>
    </queries>
'''


def patch_manifest(path):
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()

    # Add <queries> if not present (goes inside <manifest>, before <application>)
    if '<queries>' not in content:
        content = content.replace('<application', QUERIES_XML + '    <application', 1)
        print(f"Added <queries> for IMAGE_CAPTURE to {path}")

    # Add FileProvider <provider> if not present (goes inside <application>)
    if 'androidx.core.content.FileProvider' not in content:
        # Insert before </application>
        content = content.replace('</application>', PROVIDER_XML + '</application>', 1)
        print(f"Added FileProvider <provider> to {path}")

    with open(path, 'w', encoding='utf-8') as f:
        f.write(content)
    print("Manifest patched successfully.")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python patch_manifest.py <path-to-AndroidManifest.xml>")
        sys.exit(1)
    patch_manifest(sys.argv[1])
