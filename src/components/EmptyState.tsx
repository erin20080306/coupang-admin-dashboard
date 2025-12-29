type Props = {
  title: string;
  description: string;
  actionLabel: string;
  onAction: () => void;
};

export default function EmptyState({ title, description, actionLabel, onAction }: Props) {
  return (
    <div className="empty">
      <div className="emptyIcon">□</div>
      <div className="emptyTitle">{title}</div>
      <div className="emptyDesc">{description}</div>
      <button className="btnPrimary" onClick={onAction}>{actionLabel}</button>
    </div>
  );
}
