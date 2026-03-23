import {
  ExtensionOptions,
  ExtensionStorage,
  Resource,
  Activity,
  FileResource,
  FolderResource,
  CourseData,
  VideoServiceResource,
} from "types"
import * as parser from "@shared/parser"
import { getMoodleBaseURL, getURLRegex } from "@shared/regexHelpers"
import logger from "@shared/logger"

async function getLastModifiedHeader(href: string, options: ExtensionOptions) {
  if (!options.detectFileUpdates) return

  const headResponse = await fetch(href, {
    method: "HEAD",
  })
  const lastModified = headResponse.headers.get("last-modified")
  return lastModified ?? undefined
}

const courseURLRegex = getURLRegex("course")

class Course {
  link: string
  HTMLDocument: Document
  name: string
  shortcut: string
  isFirstScan: boolean
  isCoursePage: boolean
  options: ExtensionOptions

  resources: Resource[]
  previousSeenResources: string[] | null

  activities: Activity[]
  previousSeenActivities: string[] | null

  lastModifiedHeaders: Record<string, string | undefined> | undefined

  sectionIndices: Record<string, number>

  constructor(link: string, HTMLDocument: Document, options: ExtensionOptions) {
    this.link = link
    this.HTMLDocument = HTMLDocument
    this.options = options
    this.name = parser.parseCourseNameFromCoursePage(HTMLDocument, options)
    this.shortcut = parser.parseCourseShortcut(HTMLDocument, options)
    this.isFirstScan = true
    this.isCoursePage = !!link.match(courseURLRegex)

    this.resources = []
    this.previousSeenResources = null

    this.activities = []
    this.previousSeenActivities = null

    this.sectionIndices = {}
  }

  private getSectionIndex(section: string): number {
    if (this.sectionIndices[section] === undefined) {
      this.sectionIndices[section] = Object.keys(this.sectionIndices).length
    }

    return this.sectionIndices[section] + 1
  }

  private addResource(resource: Resource): void {
    if (this.previousSeenResources !== null) {
      const hasNotBeenSeenBefore = !this.previousSeenResources.includes(resource.href)
      if (hasNotBeenSeenBefore) {
        resource.isNew = true
        logger.debug(resource, "New resource detected")
      }

      if (this.options.detectFileUpdates) {
        const hasBeenUpdated =
          (this.lastModifiedHeaders ?? {})[resource.href] !== resource.lastModified
        if (!resource.isNew && hasBeenUpdated) {
          resource.isUpdated = true
        }
      }
    } else {
      // If course has never been scanned previousSeenResources don't exist
      // Never treat a resource as new when the course is scanned for the first time
      // because we're capturing the initial state of the course
      resource.isNew = false
      resource.isUpdated = false
    }

    this.resources.push(resource)
  }

  private async addFile(node: HTMLElement) {
    const href = parser.parseURLFromNode(node, "file", this.options)
    if (href === "") return

    const section = parser.parseSectionName(node, this.HTMLDocument, this.options)
    const sectionIndex = this.getSectionIndex(section)
    const resource: FileResource = {
      href,
      name: parser.parseFileNameFromNode(node),
      section,
      type: "file",
      isNew: false,
      isUpdated: false,
      resourceIndex: this.resources.length + 1,
      sectionIndex,
      lastModified: await getLastModifiedHeader(href, this.options),
    }

    this.addResource(resource)
  }

  private async addPluginFile(node: HTMLElement, partOfFolder = "") {
    let href = parser.parseURLFromNode(node, "pluginfile", this.options)
    if (href === "") return

    // Avoid duplicates
    const detectedURLs = this.resources.map((r) => r.href)
    if (detectedURLs.includes(href)) return

    const section = parser.parseSectionName(node, this.HTMLDocument, this.options)
    const sectionIndex = this.getSectionIndex(section)
    const resource: FileResource = {
      href,
      name: parser.parseFileNameFromPluginFileURL(href),
      section,
      type: "pluginfile",
      partOfFolder,
      isNew: false,
      isUpdated: false,
      resourceIndex: this.resources.length + 1,
      sectionIndex,
      lastModified: await getLastModifiedHeader(href, this.options),
    }

    this.addResource(resource)
  }

  private async addURLNode(node: HTMLElement) {
    // Make sure URL is a downloadable file
    const activityIcon: HTMLImageElement | null = node.querySelector("img.activityicon")
    if (activityIcon) {
      const imgName = activityIcon.src.split("/").pop()
      if (imgName) {
        // "icon" image is usually used for websites but I can't download full websites
        // Only support external URLs when they point to a file
        const isFile = imgName !== "icon"
        if (isFile) {
          // File has been identified as downloadable
          const href = parser.parseURLFromNode(node, "url", this.options)
          if (href === "") return

          const section = parser.parseSectionName(node, this.HTMLDocument, this.options)
          const sectionIndex = this.getSectionIndex(section)
          const resourceNode: FileResource = {
            href,
            name: parser.parseFileNameFromNode(node),
            section,
            type: "url",
            isNew: false,
            isUpdated: false,
            resourceIndex: this.resources.length + 1,
            sectionIndex,
            lastModified: await getLastModifiedHeader(href, this.options),
          }

          this.addResource(resourceNode)
        }
      }
    }
  }

  private async addFolder(node: HTMLElement) {
    const href = parser.parseURLFromNode(node, "folder", this.options)

    const section = parser.parseSectionName(node, this.HTMLDocument, this.options)
    const sectionIndex = this.getSectionIndex(section)
    const resource: FolderResource = {
      href,
      name: parser.parseFileNameFromNode(node),
      section,
      type: "folder",
      isInline: false,
      isNew: false,
      isUpdated: false,
      resourceIndex: this.resources.length + 1,
      sectionIndex,
    }

    if (resource.href === "") {
      // Folder could be displayed inline
      const downloadButtonVisible = parser.getDownloadButton(node) !== null
      const { downloadFolderAsZip } = this.options

      if (downloadFolderAsZip && downloadButtonVisible) {
        const downloadIdTag = parser.getDownloadIdTag(node)
        if (downloadIdTag === null) return

        const baseURL = getMoodleBaseURL(this.link)
        const downloadId = downloadIdTag.getAttribute("value")
        const downloadURL = `${baseURL}/mod/folder/download_folder.php?id=${downloadId}`

        resource.href = downloadURL
        resource.isInline = true
      } else {
        // Not downloading via button as ZIP
        // Download folder as individual pluginfiles
        // Look for any pluginfiles
        const folderFiles = node.querySelectorAll<HTMLElement>(
          parser.getQuerySelector("pluginfile", this.options)
        )
        for (const pluginFile of Array.from(folderFiles)) {
          await this.addPluginFile(pluginFile, resource.name)
        }
        return
      }
    }

    if (resource.href !== "") {
      resource.lastModified = await getLastModifiedHeader(resource.href, this.options)
    }

    this.addResource(resource)
  }

  private async addActivity(node: HTMLElement) {
    if (!this.isCoursePage) {
      return
    }

    const section = parser.parseSectionName(node, this.HTMLDocument, this.options)
    const sectionIndex = this.getSectionIndex(section)
    const href = parser.parseURLFromNode(node, "activity", this.options)
    if (href === "") return

    const activity: Activity = {
      href,
      name: parser.parseActivityNameFromNode(node),
      section: parser.parseSectionName(node, this.HTMLDocument, this.options),
      isNew: false,
      isUpdated: false,
      type: "activity",
      activityType: parser.parseActivityTypeFromNode(node),
      resourceIndex: this.activities.length + 1,
      sectionIndex,
    }

    if (
      this.previousSeenActivities !== null &&
      !this.previousSeenActivities.includes(activity.href)
    ) {
      activity.isNew = true
    }

    this.activities.push(activity)
  }

  private async fetchDocument(url: string): Promise<Document | null> {
    try {
      const response = await fetch(url)
      const body = await response.text()
      return new DOMParser().parseFromString(body, "text/html")
    } catch (err) {
      logger.error(err)
      return null
    }
  }

  private getMoodleSesskey(document: Document): string | null {
    const scriptContent = Array.from(document.scripts)
      .map((script) => script.textContent ?? "")
      .find((content) => content.includes('"sesskey"') || content.includes("sesskey"))

    if (!scriptContent) {
      return null
    }

    const sesskeyMatch = scriptContent.match(/"sesskey":"([^"]+)"/)
    return sesskeyMatch?.[1] ?? null
  }

  private addVideoServiceResource(resource: VideoServiceResource): void {
    const detectedURLs = this.resources.map((r) => r.href)
    if (detectedURLs.includes(resource.href)) {
      return
    }

    this.addResource(resource)
  }

  private async addVideoServiceVideo(
    videoPageURL: string,
    section: string,
    fallbackName: string,
    partOfFolder = ""
  ): Promise<void> {
    const videoDocument = await this.fetchDocument(videoPageURL)
    if (!videoDocument) return

    const videoElement = videoDocument.querySelector<HTMLVideoElement>(
      parser.getQuerySelector("videoservice", this.options)
    )
    const src = videoElement?.src
    if (!src) return

    const videoName =
      videoDocument.querySelector(".page-header-headings h1")?.textContent?.trim() ||
      videoDocument.querySelector(".heading")?.textContent?.trim() ||
      videoDocument.querySelector(".instancename")?.textContent?.trim() ||
      fallbackName

    const sectionIndex = this.getSectionIndex(section)
    const resource: VideoServiceResource = {
      href: videoPageURL,
      src,
      name: videoName,
      section,
      partOfFolder,
      type: "videoservice",
      isNew: false,
      isUpdated: false,
      resourceIndex: this.resources.length + 1,
      sectionIndex,
    }

    this.addVideoServiceResource(resource)
  }

  private isVideoServiceVideoURL(url: string): boolean {
    try {
      const parsedURL = new URL(url, location.href)
      return (
        parsedURL.pathname.includes("/mod/videoservice/view.php/") &&
        parsedURL.pathname.includes("/video/")
      )
    } catch {
      return false
    }
  }

  private isVideoServiceBrowseURL(url: string): boolean {
    try {
      const parsedURL = new URL(url, location.href)
      return (
        parsedURL.pathname.includes("/mod/videoservice/view.php") &&
        parsedURL.pathname.endsWith("/browse")
      )
    } catch {
      return false
    }
  }

  private getVideoServiceBrowseURL(url: string): string | null {
    try {
      const parsedURL = new URL(url, location.href)
      const baseURL = getMoodleBaseURL(parsedURL.href)

      const queryId = parsedURL.searchParams.get("id")
      if (queryId) {
        return `${baseURL}/mod/videoservice/view.php/cm/${queryId}/browse`
      }

      const cmPathMatch = parsedURL.pathname.match(/\/cm\/(\d+)(?:\/|$)/)
      if (cmPathMatch?.[1]) {
        return `${baseURL}/mod/videoservice/view.php/cm/${cmPathMatch[1]}/browse`
      }
    } catch {
      return null
    }

    return null
  }

  private getVideoServiceBrowseURLFromDocument(document: Document): string | null {
    const courseModuleId = this.getVideoServiceCourseModuleIdFromDocument(document)
    if (!courseModuleId) {
      return null
    }

    const canonicalHref =
      document.querySelector<HTMLLinkElement>('link[rel="canonical"]')?.href ??
      document.location?.href ??
      ""
    const baseURL = getMoodleBaseURL(canonicalHref)
    if (!baseURL) {
      return null
    }

    return `${baseURL}/mod/videoservice/view.php/cm/${courseModuleId}/browse`
  }

  private getVideoServiceCourseModuleIdFromDocument(document: Document): string | null {
    const bodyClass = document.body?.className ?? ""
    const cmClassMatch = bodyClass.match(/\bcmid-(\d+)\b/)
    return cmClassMatch?.[1] ?? null
  }

  private getVideoServiceCourseModuleId(url: string, document?: Document): string | null {
    try {
      const parsedURL = new URL(url, location.href)
      const queryId = parsedURL.searchParams.get("id")
      if (queryId) {
        return queryId
      }

      const cmPathMatch = parsedURL.pathname.match(/\/cm\/(\d+)(?:\/|$)/)
      if (cmPathMatch?.[1]) {
        return cmPathMatch[1]
      }
    } catch {
      // ignore and try document fallback below
    }

    if (!document) {
      return null
    }

    return this.getVideoServiceCourseModuleIdFromDocument(document)
  }

  private async getVideoServiceVideosFromAPI(
    courseModuleId: string
  ): Promise<
    Array<{
      id: number
      title: string
      url: string
    }>
  > {
    const sesskey = this.getMoodleSesskey(this.HTMLDocument)
    if (!sesskey) {
      return []
    }

    try {
      const response = await fetch(
        `/lib/ajax/service.php?sesskey=${encodeURIComponent(sesskey)}&info=mod_videoservice_get_videos`,
        {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify([
            {
              index: 0,
              methodname: "mod_videoservice_get_videos",
              args: {
                coursemoduleid: courseModuleId,
                courseid: 0,
              },
            },
          ]),
        }
      )
      const data = await response.json()
      const videos = data?.[0]?.data?.videos

      if (!Array.isArray(videos)) {
        return []
      }

      return videos
        .filter((video) => typeof video?.url === "string" && typeof video?.title === "string")
        .map((video) => ({
          id: video.id,
          title: video.title,
          url: video.url,
        }))
    } catch (err) {
      logger.error(err)
      return []
    }
  }

  private getVideoServiceVideoLinks(document: Document): HTMLAnchorElement[] {
    return Array.from(
      document.querySelectorAll<HTMLAnchorElement>('a[href*="/mod/videoservice/view.php"]')
    )
      .filter((link) => this.isVideoServiceVideoURL(link.href))
      .reduce((nodes, current) => {
        if (!nodes.some((node) => node.href === current.href)) {
          nodes.push(current)
        }
        return nodes
      }, [] as HTMLAnchorElement[])
  }

  private async addVideoService(node: HTMLElement) {
    const href = parser.parseURLFromNode(node, "activity", this.options)
    if (href === "") return

    const section = parser.parseSectionName(node, this.HTMLDocument, this.options)
    const fallbackName = parser.parseActivityNameFromNode(node)
    const activityDocument = await this.fetchDocument(href)
    if (!activityDocument) return
    const courseModuleId = this.getVideoServiceCourseModuleId(href, activityDocument)

    const embeddedVideo = activityDocument.querySelector<HTMLVideoElement>(
      parser.getQuerySelector("videoservice", this.options)
    )
    if (embeddedVideo?.src) {
      await this.addVideoServiceVideo(href, section, fallbackName, fallbackName)
      return
    }

    const videoLinks = this.getVideoServiceVideoLinks(activityDocument)

    if (videoLinks.length > 0) {
      await Promise.all(
        videoLinks.map((link) => {
          const videoName = link.textContent?.trim() || fallbackName
          return this.addVideoServiceVideo(link.href, section, videoName, fallbackName)
        })
      )
      return
    }

    const browseLinks = Array.from(
      activityDocument.querySelectorAll<HTMLAnchorElement>('a[href*="/mod/videoservice/view.php"]')
    ).filter((link) => this.isVideoServiceBrowseURL(link.href))

    for (const browseLink of browseLinks) {
      const browseDocument = await this.fetchDocument(browseLink.href)
      if (!browseDocument) continue

      const browseVideoLinks = this.getVideoServiceVideoLinks(browseDocument)
      if (browseVideoLinks.length === 0) continue

      await Promise.all(
        browseVideoLinks.map((link) => {
          const videoName = link.textContent?.trim() || fallbackName
          return this.addVideoServiceVideo(link.href, section, videoName, fallbackName)
        })
      )
      return
    }

    const derivedBrowseURLs = [
      this.getVideoServiceBrowseURL(href),
      this.getVideoServiceBrowseURL(activityDocument.location?.href ?? ""),
      this.getVideoServiceBrowseURLFromDocument(activityDocument),
    ].filter((value, index, array): value is string => Boolean(value) && array.indexOf(value) === index)

    for (const browseURL of derivedBrowseURLs) {
      const browseDocument = await this.fetchDocument(browseURL)
      if (!browseDocument) continue

      const browseVideoLinks = this.getVideoServiceVideoLinks(browseDocument)
      if (browseVideoLinks.length === 0) continue

      await Promise.all(
        browseVideoLinks.map((link) => {
          const videoName = link.textContent?.trim() || fallbackName
          return this.addVideoServiceVideo(link.href, section, videoName)
        })
      )
      return
    }

    if (courseModuleId) {
      const apiVideos = await this.getVideoServiceVideosFromAPI(courseModuleId)
      if (apiVideos.length > 0) {
        await Promise.all(
          apiVideos.map((video) => {
            const sectionIndex = this.getSectionIndex(section)
            const resource: VideoServiceResource = {
              href: `${location.origin}/mod/videoservice/view.php/cm/${courseModuleId}/video/${video.id}/view`,
              src: video.url,
              name: video.title || fallbackName,
              section,
              partOfFolder: fallbackName,
              type: "videoservice",
              isNew: false,
              isUpdated: false,
              resourceIndex: this.resources.length + 1,
              sectionIndex,
            }

            this.addVideoServiceResource(resource)
          })
        )
        return
      }
    }

    await this.addVideoServiceVideo(href, section, fallbackName, fallbackName)
  }

  async scan(testLocalStorage?: ExtensionStorage): Promise<void> {
    this.resources = []
    this.previousSeenResources = null

    this.activities = []
    this.previousSeenActivities = null

    this.sectionIndices = {}

    //  Local storage course data
    const localStorage =
      testLocalStorage ?? ((await chrome.storage.local.get()) as ExtensionStorage)
    const { options, courseData } = localStorage

    this.options = options

    if (courseData[this.link]) {
      // Course exists in locally stored data
      this.isFirstScan = false
      const storedCourseData = courseData[this.link]
      logger.debug(storedCourseData, "Course was found in local storage")

      this.previousSeenResources = storedCourseData.seenResources
      this.previousSeenActivities = storedCourseData.seenActivities
      this.lastModifiedHeaders = storedCourseData.lastModifiedHeaders
    } else {
      logger.debug(`New course detected ${this.name}`)
    }

    const mainHTML = this.HTMLDocument.querySelector("#region-main")

    if (!mainHTML) {
      return
    }

    const modules = mainHTML.querySelectorAll<HTMLElement>("li[id^='module-']")
    if (modules && modules.length !== 0) {
      for (const node of Array.from(modules)) {
        const isFile = node.classList.contains("resource")
        const isFolder = node.classList.contains("folder")
        const isURL = node.classList.contains("url")
        const isVideoService = node.classList.contains("videoservice")

        if (isFile) {
          await this.addFile(node)
        } else if (isFolder) {
          await this.addFolder(node)
        } else if (isURL) {
          await this.addURLNode(node)
        } else if (isVideoService) {
          await this.addVideoService(node)
        } else {
          await this.addActivity(node)
        }
      }

      // Check for pluginfiles that could be anywhere on the page
      const pluginFileNodes = Array.from(
        mainHTML.querySelectorAll<HTMLElement>(parser.getQuerySelector("pluginfile", this.options))
      )
      const mediaFileNodes = Array.from(
        mainHTML.querySelectorAll<HTMLElement>(parser.getQuerySelector("media", this.options))
      )
      await Promise.all(pluginFileNodes.map((n) => this.addPluginFile(n)))
      await Promise.all(mediaFileNodes.map((n) => this.addPluginFile(n)))
    } else {
      // Backup solution that is a little more brute force
      const fileNodes = Array.from(
        mainHTML.querySelectorAll<HTMLElement>(parser.getQuerySelector("file", this.options))
      )
      const pluginFileNodes = Array.from(
        mainHTML.querySelectorAll<HTMLElement>(parser.getQuerySelector("pluginfile", this.options))
      )
      const urlFileNodes = Array.from(
        mainHTML.querySelectorAll<HTMLElement>(parser.getQuerySelector("url", this.options))
      )
      const mediaFileNodes = Array.from(
        mainHTML.querySelectorAll<HTMLElement>(parser.getQuerySelector("media", this.options))
      )
      const folderNodes = Array.from(
        mainHTML.querySelectorAll<HTMLElement>(parser.getQuerySelector("folder", this.options))
      )
      const activities = Array.from(
        mainHTML.querySelectorAll<HTMLElement>(parser.getQuerySelector("activity", this.options))
      )

      await Promise.all(fileNodes.map((n) => this.addFile(n)))
      await Promise.all(pluginFileNodes.map((n) => this.addPluginFile(n)))
      await Promise.all(urlFileNodes.map((n) => this.addURLNode(n)))
      await Promise.all(mediaFileNodes.map((n) => this.addPluginFile(n)))
      await Promise.all(folderNodes.map((n) => this.addFolder(n)))
      await Promise.all(activities.map((n) => this.addActivity(n)))
    }

    logger.debug("Course scan finished", { course: this })

    if (testLocalStorage) {
      return
    }

    if (this.lastModifiedHeaders === undefined) {
      this.lastModifiedHeaders = Object.fromEntries(
        this.resources.map((r) => [r.href, r.lastModified])
      )
    }

    const updatedCourseData = {
      seenResources: this.resources.filter((n) => !n.isNew).map((n) => n.href),
      newResources: this.resources.filter((n) => n.isNew).map((n) => n.href),
      seenActivities: this.activities.filter((n) => !n.isNew).map((n) => n.href),
      newActivities: this.activities.filter((n) => n.isNew).map((n) => n.href),
      lastModifiedHeaders: this.lastModifiedHeaders,
    }
    courseData[this.link] = updatedCourseData

    logger.debug(`Storing course data in local storage for course ${this.name}`, {
      updatedCourseData,
    })
    await chrome.storage.local.set({ courseData } satisfies Partial<ExtensionStorage>)
  }

  async updateStoredResources(downloadedResources?: Resource[]): Promise<CourseData> {
    const { courseData } = (await chrome.storage.local.get("courseData")) as ExtensionStorage
    const storedCourseData = courseData[this.link]
    const { seenResources, lastModifiedHeaders } = storedCourseData

    const newResources = this.resources.filter((n) => n.isNew)

    // Default behavior: Merge all stored new resources
    let toBeMerged = newResources

    // If downloaded resources are provided then only merge those
    if (downloadedResources) {
      toBeMerged = downloadedResources
    }

    // Merge already seen resources with new resources
    // Use set to remove duplicates
    logger.debug(toBeMerged, "Adding resources to list of seen resources")
    const updatedSeenResources = Array.from(
      new Set(seenResources.concat(toBeMerged.map((r) => r.href)))
    )

    const updatedNewResources = newResources
      .filter((r) => !updatedSeenResources.includes(r.href))
      .map((r) => r.href)

    if (lastModifiedHeaders) {
      const toBeUpdated = toBeMerged

      if (downloadedResources === undefined) {
        const updatedResources = this.resources.filter((n) => n.isUpdated)
        toBeUpdated.push(...updatedResources)
      }

      toBeUpdated.forEach((r) => {
        lastModifiedHeaders[r.href] = r.lastModified
        r.isNew = false
        r.isUpdated = false
      })
    }

    const updatedCourseData = {
      ...(storedCourseData as CourseData),
      seenResources: updatedSeenResources,
      newResources: updatedNewResources,
      lastModifiedHeaders,
    } satisfies CourseData

    logger.debug(updatedCourseData, "Storing updated course data in local storage")
    await chrome.storage.local.set({
      courseData: {
        ...courseData,
        [this.link]: updatedCourseData,
      },
    } satisfies Partial<ExtensionStorage>)

    return updatedCourseData
  }

  async updateStoredActivities(): Promise<CourseData> {
    const { courseData } = (await chrome.storage.local.get("courseData")) as ExtensionStorage
    const storedCourseData = courseData[this.link]

    const { seenActivities, newActivities } = storedCourseData
    logger.debug(newActivities, "Adding activities to list of seen activities")
    const updatedSeenActivities = Array.from(new Set(seenActivities.concat(newActivities)))
    const updatedNewActivities: string[] = []

    const updatedCourseData = {
      ...(storedCourseData as CourseData),
      seenActivities: updatedSeenActivities,
      newActivities: updatedNewActivities,
    } satisfies CourseData

    await chrome.storage.local.set({
      courseData: {
        ...courseData,
        [this.link]: updatedCourseData,
      },
    } satisfies Partial<ExtensionStorage>)

    return updatedCourseData
  }

  getNumberOfUpdates(): number {
    return [...this.resources, ...this.activities].filter((r) => r.isNew || r.isUpdated).length
  }
}

export default Course
