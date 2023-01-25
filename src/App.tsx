import React from "react";
import "./App.css";

const host = "https://jsonplaceholder.typicode.com/todos";

interface PlaceholderDataTypes {
  userId: number;
  id: number;
  title: string;
  completed: boolean;
}

interface PlaceholderListDataTypes {
  data: PlaceholderDataTypes[];
}

const fetchData = async (
  abort: AbortController
): Promise<PlaceholderDataTypes[]> => {
  try {
    const res = await fetch(host, { signal: abort.signal });
    const resData = (await res.json()) as PlaceholderDataTypes[];
    return resData;
  } catch (error: any) {
    throw new Error(error);
  }
};

const Posts = (data: PlaceholderDataTypes) => (
  <div style={{ margin: 2 }}>
    <h1>{data.id}</h1>
    <h1>{data.title}</h1>
    <div
      style={{
        background: "white",
        width: "100%",
        height: "1px",
      }}
    />
  </div>
);

const ListPost: React.FC<PlaceholderListDataTypes> = ({ data }) => (
  <React.Fragment>
    {data.map((it, idx) => (
      <Posts key={idx} {...it} />
    ))}
  </React.Fragment>
);

const Lists = React.memo(ListPost);

function App() {
  const [isLoading, setIsLoading] = React.useState(true);
  const [data, setData] = React.useState<PlaceholderDataTypes[]>([]);
  const [datas, setDatas] = React.useState<PlaceholderDataTypes[]>([]);

  React.useEffect(() => {
    const abort = new AbortController();

    const getData = async () => {
      const data = await fetchData(abort);
      setData(data);
      setIsLoading(false);
    };
    getData();

    return () => abort.abort();
  }, []);

  const setFilter = React.useMemo(
    () => (e: string) => {
      if (!e) return setDatas([]);
      const filterData = data.filter((it) => {
        return it.title.includes(e);
      });
      setDatas(filterData);
    },
    [data]
  );

  return (
    <div className="App">
      {isLoading ? (
        <p>Loading...</p>
      ) : (
        <div className="box">
          <input
            className="input"
            onChange={(e) => setFilter(e.target.value)}
          />
          <Lists data={datas.length > 0 ? datas : data} />
        </div>
      )}
    </div>
  );
}

export default App;
