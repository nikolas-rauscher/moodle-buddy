cd ..

# Get version
echo "Reading version from manifest.json..."
EXT_VERSION=v$(cat build/manifest.json | grep \"version\" | head -1 | awk -F: '{ print $2 }' | sed 's/[",]//g' | tr -d '[[:space:]]')
echo "Detected manifest version $EXT_VERSION"

LATEST_TAG=$(git describe --tags $(git rev-list --tags --max-count=1))
echo "Latest git tag: $LATEST_TAG"
# if the tags are equal throw error
if [ "$EXT_VERSION" == "$LATEST_TAG" ]; then
  echo "Version $EXT_VERSION already exists."
fi

# Check if extension was built
if [ ! -f "moodle-buddy-firefox.zip" ]; then
  echo "moodle-buddy-firefox.zip does not exist."
  exit
fi
if [ ! -f "moodle-buddy-chrome.zip" ]; then
  echo "moodle-buddy-chrome.zip does not exist."
  exit
fi

mkdir tmp
cp moodle-buddy-firefox.zip "tmp/moodle-buddy-firefox-$EXT_VERSION.zip"
cp moodle-buddy-chrome.zip "tmp/moodle-buddy-chrome-$EXT_VERSION.zip"

RELEASE_NOTES_FILE="release-notes/$EXT_VERSION.md"
if [ ! -f "$RELEASE_NOTES_FILE" ]; then
  echo "Release notes file $RELEASE_NOTES_FILE does not exist."
  exit 1
fi

gh release create "$EXT_VERSION" --notes-file "$RELEASE_NOTES_FILE" --title "Version $EXT_VERSION" "tmp/moodle-buddy-firefox-$EXT_VERSION.zip" "tmp/moodle-buddy-chrome-$EXT_VERSION.zip"

rm -rf tmp
