import * as React from 'react';
import { useState, useEffect, useCallback } from 'react';
import { SitePagesService, IFolder } from '../services/SitePagesService';
import styles from './MoveFolderPicker.module.scss';

export interface IMoveFolderPickerProps {
  service: SitePagesService;
  libraryTitle: string;
  selectedFiles: { name: string; serverRelativeUrl: string }[];
  onDismiss: () => void;
  onMoveComplete: () => void;
}

interface IBreadcrumb {
  name: string;
  serverRelativeUrl: string;
}

export const MoveFolderPicker: React.FC<IMoveFolderPickerProps> = ({
  service,
  libraryTitle,
  selectedFiles,
  onDismiss,
  onMoveComplete
}) => {
  const [currentPath, setCurrentPath] = useState<string>(service.libraryRootUrl);
  const [breadcrumbs, setBreadcrumbs] = useState<IBreadcrumb[]>([{ name: libraryTitle, serverRelativeUrl: service.libraryRootUrl }]);
  const [folders, setFolders] = useState<IFolder[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [moving, setMoving] = useState<boolean>(false);
  const [error, setError] = useState<string | null>(null);
  const [movedAny, setMovedAny] = useState<boolean>(false);

  // Files not already in the folder currently being viewed
  const filesToMove = selectedFiles.filter(file => {
    const parentFolder = file.serverRelativeUrl.substring(0, file.serverRelativeUrl.lastIndexOf('/'));
    return parentFolder.toLowerCase() !== currentPath.toLowerCase();
  });

  // After a partial move the list is out of date, so closing must refresh it
  const handleDismiss = movedAny ? onMoveComplete : onDismiss;

  const loadFolders = useCallback(async (path: string): Promise<void> => {
    setLoading(true);
    setError(null);
    try {
      const result = await service.getFolders(path);
      setFolders(result);
    } catch {
      setError('Failed to load folders. Please try again.');
    } finally {
      setLoading(false);
    }
  }, [service]);

  useEffect(() => {
    loadFolders(currentPath).catch(() => {/* handled inside */});
  }, [currentPath, loadFolders]);

  const navigateInto = (folder: IFolder): void => {
    setCurrentPath(folder.ServerRelativeUrl);
    setBreadcrumbs(prev => [...prev, { name: folder.Name, serverRelativeUrl: folder.ServerRelativeUrl }]);
  };

  const navigateTo = (index: number): void => {
    const crumb = breadcrumbs[index];
    setCurrentPath(crumb.serverRelativeUrl);
    setBreadcrumbs(prev => prev.slice(0, index + 1));
  };

  const handleMove = async (): Promise<void> => {
    setMoving(true);
    setError(null);
    const failed: string[] = [];

    for (const file of filesToMove) {
      try {
        await service.movePage(file.serverRelativeUrl, currentPath, file.name);
        setMovedAny(true);
      } catch (e) {
        failed.push(`${file.name} (${(e as Error).message})`);
      }
    }

    setMoving(false);

    if (failed.length === 0) {
      onMoveComplete();
    } else {
      setError(`Failed to move: ${failed.join('; ')}`);
    }
  };

  return (
    <div className={styles.overlay}>
      <div className={styles.dialog}>
        <div className={styles.header}>
          <span className={styles.title}>Move {selectedFiles.length > 1 ? `${selectedFiles.length} pages` : `"${selectedFiles[0].name}"`}</span>
          <button className={styles.closeButton} onClick={handleDismiss} aria-label="Close">✕</button>
        </div>

        <div className={styles.warning}>
          ⚠️ Moving a page changes its URL. Any navigation links to this page will need updating.
        </div>

        <div className={styles.breadcrumb}>
          {breadcrumbs.map((crumb, i) => (
            <React.Fragment key={crumb.serverRelativeUrl}>
              {i > 0 && <span className={styles.separator}>›</span>}
              <button
                className={`${styles.crumbButton} ${i === breadcrumbs.length - 1 ? styles.crumbActive : ''}`}
                onClick={() => navigateTo(i)}
                disabled={i === breadcrumbs.length - 1}
              >
                {crumb.name}
              </button>
            </React.Fragment>
          ))}
        </div>

        <div className={styles.folderList}>
          {loading && <div className={styles.loading}>Loading folders...</div>}
          {!loading && error && <div className={styles.error}>{error}</div>}
          {!loading && !error && folders.length === 0 && (
            <div className={styles.empty}>No subfolders here.</div>
          )}
          {!loading && folders.map(folder => (
            <button
              key={folder.ServerRelativeUrl}
              className={styles.folderItem}
              onClick={() => navigateInto(folder)}
            >
              <span className={styles.folderIcon}>📁</span>
              <span>{folder.Name}</span>
              <span className={styles.chevron}>›</span>
            </button>
          ))}
        </div>

        <div className={styles.destination}>
          Moving to: <strong>{breadcrumbs[breadcrumbs.length - 1].name}</strong>
          {filesToMove.length === 0 && ' (already in this folder)'}
        </div>

        <div className={styles.footer}>
          <button className={styles.cancelButton} onClick={handleDismiss} disabled={moving}>{movedAny ? 'Close' : 'Cancel'}</button>
          <button className={styles.moveButton} onClick={handleMove} disabled={moving || loading || filesToMove.length === 0}>
            {moving ? 'Moving...' : 'Move here'}
          </button>
        </div>
      </div>
    </div>
  );
};