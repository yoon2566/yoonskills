# 3D, Media, And Browser Workflows

## 3D Assets

### OpenSCAD

- Check that `openscad` is installed before relying on it.
- Keep preview quality lower than final export quality.
- Validate dimensions, wall thickness, manifold status, and output size.
- Do not treat a fixed `$fn` value or file-size limit as universally correct.

### Blender

- Check `blender --help` and the installed Blender version before using version-sensitive operators.
- Use a Python script file with `blender --background --python <script.py>`, or use `--python-expr` for a short expression.
- Do not use `blender --background --python -c ...`; `--python` expects a script path.
- Verify the generated or exported file.

### STL To STEP/STP

- Prefer a dedicated local conversion skill or deterministic converter.
- Preserve `.step` versus `.stp` according to the user's request.
- Explain that faceted STL conversion does not recreate smooth parametric CAD features.
- Do not upload private or proprietary models to an online converter without explicit approval.

## Browser Verification

- Use the configured browser skill or plugin.
- After changing local HTML or a web application, open the relevant target and verify it.
- Check visible layout, navigation, interactive behavior, and console errors.
- Avoid automating authenticated sites beyond the user's request.

## Video And Audio

- Inspect available media tools before selecting a workflow.
- Use a dedicated video-analysis or Remotion skill when available.
- Keep source media unchanged and write outputs to new files.
- Verify duration, audio presence, resolution, and output playback.
- Confirm rights and user intent before downloading third-party media.
