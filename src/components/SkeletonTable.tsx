export default function SkeletonTable() {
  return (
    <div className="skeletonWrap">
      <div className="skeletonHead">
        <div className="sk skText" />
        <div className="sk skText" />
        <div className="sk skText" />
      </div>
      <div className="skeletonBody">
        {Array.from({ length: 10 }).map((_, i) => (
          <div className="sk skRow" key={i} />
        ))}
      </div>
    </div>
  );
}
