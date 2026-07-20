/**
 * Public compatibility facade for the canonical PPTX slide-directory helper.
 * Keep imports from this path working while all versioning and resolution
 * semantics live in one module.
 */
export {
  assertSlideDirectoryVersion,
  createSlideDirectorySnapshot,
  getSlideDirectoryVersion,
  loadSlideDirectory,
  type ResolvedSlideTarget,
  resolveSlideIds,
  resolveSlideTarget,
  SlideDirectoryChangedDuringWriteError,
  SlideDirectoryChangedError,
  type SlideDirectorySnapshot,
  type SlideMutationError,
  SlideMutationNotStartedError,
  type SlideMutationState,
  SlideMutationUncertainError,
  SlideTargetNotFoundError,
  type SlideTargetReference,
  slideTargetParameterProperties,
  slideTargetParameters,
  toSlideMutationNotStartedError,
  toSlideMutationUncertainError,
  toSlideTargetReference,
} from "./pptx/slide-directory";
