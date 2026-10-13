import * as React from 'react';
import { createPortal } from 'react-dom';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import styles from './FTLApp.module.scss';

interface IFavoriteToastProps {
  saving: boolean;
  error?: string;
  success: boolean;
  onRetry: () => void;
  onDismiss: () => void;
  inline?: boolean;
}

export default function FavoriteToast(props: IFavoriteToastProps): React.ReactElement {
  if (!props.saving && !props.error && !props.success) return <></>;
  const notification = (
    <div className={`${styles.favoriteToast} ${props.error ? styles.toastError : ''}`}
      role={props.error ? 'alert' : 'status'} aria-live={props.error ? 'assertive' : 'polite'} aria-atomic="true">
      {props.saving && <Spinner size={SpinnerSize.small} styles={{ circle: { borderColor: '#475569', borderTopColor: '#a5b4fc' } }} />}
      <div className={styles.toastMessage}>
        {props.error || (props.saving ? 'お気に入りを保存中…' : 'お気に入りを保存しました')}
      </div>
      {props.error && <button type="button" onClick={props.onRetry}>再試行</button>}
      {props.success && !props.error && !props.saving && <button type="button" aria-label="通知を閉じる" onClick={props.onDismiss}>閉じる</button>}
    </div>
  );
  return props.inline ? notification : createPortal(notification, document.body);
}
