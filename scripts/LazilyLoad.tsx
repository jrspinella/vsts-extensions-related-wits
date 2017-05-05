import * as React from "react";

interface ILazilyLoadProps {
    modules: string[];
}

interface ILazilyLoadState {
    isLoaded: boolean;
    fetchedModules: any[];
}

export class LazilyLoad extends React.Component<ILazilyLoadProps, ILazilyLoadState> {
    private _isMounted: boolean;

    constructor(props: ILazilyLoadProps, context?: any) {
        super(props, context);

        this._isMounted = false;
        this.state = {
            isLoaded: false,
            fetchedModules: null
        };        
    }

    public componentDidMount() {
        this._isMounted = true;
        this._load();
    }

    public componentDidUpdate(previousProps: ILazilyLoadProps) {
        const shouldLoad = !!Object.keys(this.props.modules).filter((key)=> {
            return this.props.modules[key] !== previous.modules[key];
        }).length;
        if (shouldLoad) {
            this._load();
        }
    }

    public componentWillUnmount() {
        this._isMounted = false;
    }

    private _load() {
        this.setState({
            isLoaded: false,
        });

        const { modules } = this.props;
        const keys = Object.keys(modules);

        Promise.all(keys.map((key) => modules[key]()))
        .then((values) => (keys.reduce((agg, key, index) => {
            agg[key] = values[index];
            return agg;
        }, {})))
        .then((result) => {
            if (!this._isMounted) return null;
            this.setState({ modules: result, isLoaded: true });
        });
    }

    render() {
        if (!this.state.isLoaded) return null;
        return React.Children.only(this.props.children(this.state.modules));
    }
}